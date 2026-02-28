import { useEffect, useState, useCallback } from 'react'
import { api } from '../api/bridge'
import type { TaskInfo } from '../types'

export default function BatchPage() {
  const [tasks, setTasks] = useState<TaskInfo[]>([])

  const refresh = useCallback(async () => {
    const bridge = api()
    if (!bridge) return
    const res = await bridge.list_tasks()
    if (res.ok && res.data) setTasks(res.data)
  }, [])

  useEffect(() => {
    refresh()
    const timer = setInterval(refresh, 2000)
    return () => clearInterval(timer)
  }, [refresh])

  const handleCancel = async (taskId: string) => {
    const bridge = api()
    if (!bridge) return
    await bridge.cancel_task(taskId)
    refresh()
  }

  const handleOpenFolder = async (path: string) => {
    const bridge = api()
    if (!bridge) return
    await bridge.open_output_folder(path)
  }

  return (
    <div>
      <h1 className="page-title">批处理任务队列</h1>

      <div style={{ display: 'flex', gap: 10, marginBottom: 16 }}>
        <button className="btn-secondary" onClick={refresh}>
          刷新
        </button>
      </div>

      {tasks.length === 0 ? (
        <div className="card" style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: 32 }}>
          暂无任务
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {tasks.map((task) => (
            <div key={task.task_id} className="card" style={{ padding: '12px 16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 6 }}>
                <span className={`status-badge ${task.status}`}>{task.status}</span>
                <span style={{ fontWeight: 500, fontSize: 14 }}>{task.task_type}</span>
                <span style={{ fontSize: 12, color: 'var(--text-secondary)' }}>{task.task_id}</span>
                <span style={{ flex: 1 }} />
                {task.status === 'running' && (
                  <button className="btn-danger" style={{ padding: '4px 10px', fontSize: 12 }} onClick={() => handleCancel(task.task_id)}>
                    取消
                  </button>
                )}
                {task.output_dir && task.status === 'success' && (
                  <button className="btn-secondary" style={{ padding: '4px 10px', fontSize: 12 }} onClick={() => handleOpenFolder(task.output_dir!)}>
                    打开目录
                  </button>
                )}
              </div>
              <div style={{ fontSize: 13, color: 'var(--text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {task.input_path}
              </div>
              {task.summary && (
                <div style={{ fontSize: 12, color: 'var(--text-secondary)', marginTop: 4 }}>
                  处理: {task.summary.processed} | 失败: {task.summary.failed} | 跳过: {task.summary.skipped}
                </div>
              )}
              {task.error && (
                <div style={{ fontSize: 12, color: 'var(--error)', marginTop: 4 }}>
                  {task.error.code}: {task.error.message}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
