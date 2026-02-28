"""任务执行器 — 串行/并行、取消、重试"""

from __future__ import annotations

import threading
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from dataclasses import dataclass
from enum import Enum
from typing import Callable

from core.api import TaskRequest, TaskResponse, run_task
from core.logging_utils import get_logger

logger = get_logger()


class JobStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    FAILED = "failed"
    CANCELLED = "cancelled"


@dataclass
class Job:
    request: TaskRequest
    status: JobStatus = JobStatus.PENDING
    response: TaskResponse | None = None
    retries: int = 0


class TaskRunner:
    """管理一组任务的串行/并行执行、取消与重试"""

    def __init__(self, max_workers: int = 1):
        self.max_workers = max(1, max_workers)
        self._jobs: list[Job] = []
        self._cancel_event = threading.Event()
        self._lock = threading.Lock()
        self._on_progress: Callable[[Job, int, int], None] | None = None

    def set_progress_callback(self, callback: Callable[[Job, int, int], None]):
        self._on_progress = callback

    def submit(self, request: TaskRequest) -> Job:
        job = Job(request=request)
        with self._lock:
            self._jobs.append(job)
        return job

    def cancel(self):
        self._cancel_event.set()

    def reset(self):
        self._cancel_event.clear()
        with self._lock:
            self._jobs.clear()

    @property
    def jobs(self) -> list[Job]:
        return list(self._jobs)

    def run_all(self) -> list[Job]:
        self._cancel_event.clear()
        total = len(self._jobs)

        if self.max_workers <= 1:
            return self._run_serial(total)
        return self._run_parallel(total)

    def retry_failed(self, max_retries: int = 1) -> list[Job]:
        """重跑所有失败的任务"""
        failed = [j for j in self._jobs if j.status == JobStatus.FAILED and j.retries < max_retries]
        for job in failed:
            job.status = JobStatus.PENDING
            job.retries += 1
        self._cancel_event.clear()
        return self._run_jobs(failed, len(failed))

    # ---- internal ----

    def _run_serial(self, total: int) -> list[Job]:
        return self._run_jobs(self._jobs, total)

    def _run_jobs(self, jobs: list[Job], total: int) -> list[Job]:
        done_count = 0
        for job in jobs:
            if self._cancel_event.is_set():
                job.status = JobStatus.CANCELLED
                done_count += 1
                self._notify(job, done_count, total)
                continue
            job.status = JobStatus.RUNNING
            try:
                job.response = run_task(job.request, cancel_event=self._cancel_event)
                job.status = self._status_from_response(job.response)
            except Exception as e:
                logger.error("任务执行异常: %s", e)
                job.status = JobStatus.FAILED
            done_count += 1
            self._notify(job, done_count, total)
        return jobs

    def _run_parallel(self, total: int) -> list[Job]:
        done_count = 0
        next_job_idx = 0

        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            future_map: dict[Future, Job] = {}

            def submit_until_full():
                nonlocal next_job_idx
                while (
                    len(future_map) < self.max_workers
                    and next_job_idx < len(self._jobs)
                    and not self._cancel_event.is_set()
                ):
                    job = self._jobs[next_job_idx]
                    next_job_idx += 1
                    job.status = JobStatus.RUNNING
                    fut = pool.submit(run_task, job.request, self._cancel_event)
                    future_map[fut] = job

            def mark_remaining_cancelled():
                nonlocal done_count, next_job_idx
                while next_job_idx < len(self._jobs):
                    job = self._jobs[next_job_idx]
                    next_job_idx += 1
                    if job.status == JobStatus.PENDING:
                        job.status = JobStatus.CANCELLED
                        done_count += 1
                        self._notify(job, done_count, total)

            submit_until_full()
            if self._cancel_event.is_set():
                mark_remaining_cancelled()

            while future_map:
                done, _ = wait(
                    list(future_map.keys()),
                    return_when=FIRST_COMPLETED,
                )
                for fut in done:
                    job = future_map.pop(fut)
                    try:
                        job.response = fut.result()
                        job.status = self._status_from_response(job.response)
                    except Exception as e:
                        logger.error("并行任务异常: %s", e)
                        job.status = (
                            JobStatus.CANCELLED
                            if self._cancel_event.is_set()
                            else JobStatus.FAILED
                        )
                    done_count += 1
                    self._notify(job, done_count, total)

                if self._cancel_event.is_set():
                    for fut, job in list(future_map.items()):
                        if fut.cancel():
                            future_map.pop(fut, None)
                            job.status = JobStatus.CANCELLED
                            done_count += 1
                            self._notify(job, done_count, total)
                    mark_remaining_cancelled()
                else:
                    submit_until_full()
        return self._jobs

    @staticmethod
    def _status_from_response(response: TaskResponse) -> JobStatus:
        if response.ok:
            return JobStatus.SUCCESS
        error_code = ""
        if response.error:
            error_code = str(response.error.get("code", ""))
        if error_code == "E_CANCELLED":
            return JobStatus.CANCELLED
        return JobStatus.FAILED

    def _notify(self, job: Job, current: int, total: int):
        if self._on_progress:
            try:
                self._on_progress(job, current, total)
            except Exception:
                pass
