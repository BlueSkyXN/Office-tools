import { useState, useCallback } from 'react'
import FilePicker from '../components/FilePicker'
import TaskControls from '../components/TaskControls'
import { api } from '../api/bridge'
import type { ImageOptions, TaskInfo } from '../types'

export default function ImageToolPage() {
  const [inputPath, setInputPath] = useState('')
  const [outputDir, setOutputDir] = useState('')
  const [options, setOptions] = useState<ImageOptions>({
    remove_images: false,
    optimize_images: false,
    jpeg_quality: 85,
  })
  const [running, setRunning] = useState(false)
  const [taskId, setTaskId] = useState<string | null>(null)
  const [result, setResult] = useState<TaskInfo | null>(null)

  const handleStart = useCallback(async () => {
    const bridge = api()
    if (!bridge || !inputPath) return
    setRunning(true)
    setResult(null)

    const res = await bridge.start_task({
      task_type: 'image_extract',
      input_path: inputPath,
      output_dir: outputDir || undefined,
      options,
    })

    if (res.ok && res.data) {
      setTaskId(res.data.task_id)
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
      <h1 className="page-title">图片分离与标记</h1>

      <div className="card" style={{ maxWidth: 600 }}>
        <FilePicker label="输入文件" value={inputPath} onChange={setInputPath} />
        <FilePicker label="输出目录（可选）" value={outputDir} onChange={setOutputDir} mode="folder" />

        <div style={{ marginTop: 16 }}>
          <label style={{ fontSize: 13, fontWeight: 500, color: 'var(--text-secondary)', marginBottom: 8, display: 'block' }}>
            处理选项
          </label>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <label className="checkbox-row">
              <input
                type="checkbox"
                checked={options.remove_images}
                onChange={() => setOptions((p) => ({ ...p, remove_images: !p.remove_images }))}
              />
              删除原图（仅保留标记）
            </label>
            <label className="checkbox-row">
              <input
                type="checkbox"
                checked={options.optimize_images}
                onChange={() => setOptions((p) => ({ ...p, optimize_images: !p.optimize_images }))}
              />
              启用图片优化
            </label>
            <div className="form-group" style={{ marginBottom: 0 }}>
              <label>JPEG 质量 ({options.jpeg_quality})</label>
              <input
                type="range"
                min={1}
                max={100}
                value={options.jpeg_quality}
                onChange={(e) => setOptions((p) => ({ ...p, jpeg_quality: Number(e.target.value) }))}
                style={{ width: '100%', border: 'none', padding: 0 }}
              />
            </div>
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
            <span className={`status-badge ${result.status}`}>{result.status}</span>
            {result.summary && (
              <div style={{ fontSize: 13, color: 'var(--text-secondary)', marginTop: 4 }}>
                处理: {result.summary.processed} | 失败: {result.summary.failed} | 跳过: {result.summary.skipped}
              </div>
            )}
            {result.error && (
              <div style={{ fontSize: 13, color: 'var(--error)', marginTop: 4 }}>{result.error.message}</div>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
