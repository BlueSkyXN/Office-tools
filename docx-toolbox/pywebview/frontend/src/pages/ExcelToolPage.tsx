import { useState, useCallback } from 'react'
import FilePicker from '../components/FilePicker'
import TaskControls from '../components/TaskControls'
import { api } from '../api/bridge'
import type { ExcelOptions, TaskInfo } from '../types'

export default function ExcelToolPage() {
  const [inputPath, setInputPath] = useState('')
  const [outputDir, setOutputDir] = useState('')
  const [options, setOptions] = useState<ExcelOptions>({
    word_table: true,
    extract_excel: false,
    image: false,
    keep_attachment: false,
    remove_watermark: false,
    a3: false,
  })
  const [running, setRunning] = useState(false)
  const [taskId, setTaskId] = useState<string | null>(null)
  const [result, setResult] = useState<TaskInfo | null>(null)

  const toggle = (key: keyof ExcelOptions) => {
    setOptions((prev) => ({ ...prev, [key]: !prev[key] }))
  }

  const handleStart = useCallback(async () => {
    const bridge = api()
    if (!bridge || !inputPath) return
    setRunning(true)
    setResult(null)

    const res = await bridge.start_task({
      task_type: 'excel_allinone',
      input_path: inputPath,
      output_dir: outputDir || undefined,
      options,
    })

    if (res.ok && res.data) {
      setTaskId(res.data.task_id)
      // Poll for completion
      const poll = setInterval(async () => {
        const status = await bridge.get_task_status(res.data!.task_id)
        if (status.ok && status.data) {
          const s = status.data.status
          if (s === 'success' || s === 'failed' || s === 'cancelled') {
            clearInterval(poll)
            setRunning(false)
            setResult(status.data)
          }
        }
      }, 1000)
    } else {
      setRunning(false)
    }
  }, [inputPath, outputDir, options])

  const handleCancel = useCallback(async () => {
    const bridge = api()
    if (!bridge || !taskId) return
    await bridge.cancel_task(taskId)
    setRunning(false)
  }, [taskId])

  return (
    <div>
      <h1 className="page-title">Excel 嵌入对象处理</h1>

      <div className="card" style={{ maxWidth: 600 }}>
        <FilePicker label="输入文件" value={inputPath} onChange={setInputPath} />
        <FilePicker label="输出目录（可选）" value={outputDir} onChange={setOutputDir} mode="folder" />

        <div style={{ marginTop: 16 }}>
          <label style={{ fontSize: 13, fontWeight: 500, color: 'var(--text-secondary)', marginBottom: 8, display: 'block' }}>
            处理选项
          </label>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            <label className="checkbox-row">
              <input type="checkbox" checked={options.word_table} onChange={() => toggle('word_table')} />
              转换为 Word 表格
            </label>
            <label className="checkbox-row">
              <input type="checkbox" checked={options.extract_excel} onChange={() => toggle('extract_excel')} />
              提取嵌入 Excel
            </label>
            <label className="checkbox-row">
              <input type="checkbox" checked={options.image} onChange={() => toggle('image')} />
              渲染为图片
            </label>
            <label className="checkbox-row">
              <input type="checkbox" checked={options.keep_attachment} onChange={() => toggle('keep_attachment')} />
              保留附件入口
            </label>
            <label className="checkbox-row">
              <input type="checkbox" checked={options.remove_watermark} onChange={() => toggle('remove_watermark')} />
              移除水印
            </label>
            <label className="checkbox-row">
              <input type="checkbox" checked={options.a3} onChange={() => toggle('a3')} />
              A3 横向
            </label>
          </div>
        </div>

        <TaskControls
          running={running}
          onStart={handleStart}
          onCancel={handleCancel}
          disabled={!inputPath}
        />

        {result && (
          <div style={{ marginTop: 16, padding: 12, borderRadius: 'var(--radius-sm)', background: result.status === 'success' ? '#F0FDF4' : '#FEF2F2' }}>
            <div style={{ fontWeight: 500, marginBottom: 4 }}>
              <span className={`status-badge ${result.status}`}>{result.status}</span>
            </div>
            {result.summary && (
              <div style={{ fontSize: 13, color: 'var(--text-secondary)' }}>
                处理: {result.summary.processed} | 失败: {result.summary.failed} | 跳过: {result.summary.skipped}
              </div>
            )}
            {result.error && (
              <div style={{ fontSize: 13, color: 'var(--error)' }}>
                {result.error.message}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
