/** 共享 TypeScript 类型 — 对齐 CORE-INTERFACE.md */

export type TaskType = 'excel_allinone' | 'image_extract' | 'table_extract'

export type TaskStatusValue = 'pending' | 'running' | 'success' | 'failed' | 'cancelled'

export interface ExcelOptions {
  word_table?: boolean
  extract_excel?: boolean
  image?: boolean
  keep_attachment?: boolean
  remove_watermark?: boolean
  a3?: boolean
}

export interface ImageOptions {
  remove_images?: boolean
  optimize_images?: boolean
  jpeg_quality?: number
}

export interface TableOptions {
  include_marked?: boolean
}

export type TaskOptions = ExcelOptions | ImageOptions | TableOptions

export interface RuntimeOptions {
  workers?: number
  dry_run?: boolean
}

export interface TaskPayload {
  task_type: TaskType
  input_path: string
  output_dir?: string | null
  options?: TaskOptions
  runtime?: RuntimeOptions
}

export interface TaskSummary {
  processed: number
  failed: number
  skipped: number
  outputs: string[]
}

export interface TaskErrorInfo {
  code: string
  message: string
  detail?: string
}

export interface TaskInfo {
  task_id: string
  task_type: TaskType
  input_path: string
  output_dir: string | null
  status: TaskStatusValue
  summary: TaskSummary | null
  error: TaskErrorInfo | null
  log_path: string | null
  created_at: string
}

export interface ApiResponse<T = unknown> {
  ok: boolean
  data?: T
  error?: { code: string; message: string }
}
