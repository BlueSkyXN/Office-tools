/** Type-safe wrapper for window.pywebview.api — 遵循 DESIGN.md §6 */

import type { ApiResponse, TaskPayload, TaskInfo } from '../types'

interface PyWebViewApi {
  select_input_path(): Promise<ApiResponse<string | null>>
  select_folder(): Promise<ApiResponse<string | null>>
  start_task(payload: TaskPayload): Promise<ApiResponse<TaskInfo>>
  cancel_task(task_id: string): Promise<ApiResponse<{ task_id: string; status: string }>>
  get_task_status(task_id: string): Promise<ApiResponse<TaskInfo>>
  list_tasks(): Promise<ApiResponse<TaskInfo[]>>
  open_output_folder(path: string): Promise<ApiResponse<null>>
  export_logs(task_id: string): Promise<ApiResponse<{ path: string }>>
}

declare global {
  interface Window {
    pywebview?: { api: PyWebViewApi }
  }
}

export const api = (): PyWebViewApi | undefined => window.pywebview?.api

/** 等待 pywebview 就绪（开发模式下可能需要等待注入） */
export function waitForApi(timeout = 5000): Promise<PyWebViewApi> {
  return new Promise((resolve, reject) => {
    if (window.pywebview?.api) {
      resolve(window.pywebview.api)
      return
    }
    const start = Date.now()
    const timer = setInterval(() => {
      if (window.pywebview?.api) {
        clearInterval(timer)
        resolve(window.pywebview.api)
      } else if (Date.now() - start > timeout) {
        clearInterval(timer)
        reject(new Error('pywebview API 未就绪'))
      }
    }, 100)
  })
}
