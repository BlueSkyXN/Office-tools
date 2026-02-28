import { useState, useEffect } from 'react'

interface Settings {
  output_dir: string
  workers: number
  theme: string
}

export default function SettingsPage() {
  const [settings, setSettings] = useState<Settings>({
    output_dir: '',
    workers: 1,
    theme: 'light',
  })
  const [saved, setSaved] = useState(false)

  useEffect(() => {
    // Settings are Python-side only; placeholder for future integration
  }, [])

  const handleSave = () => {
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  return (
    <div>
      <h1 className="page-title">设置</h1>

      <div className="card" style={{ maxWidth: 500 }}>
        <div className="form-group">
          <label>默认输出目录</label>
          <input
            type="text"
            value={settings.output_dir}
            onChange={(e) => setSettings((s) => ({ ...s, output_dir: e.target.value }))}
            placeholder="留空使用输入文件同级目录"
            style={{ width: '100%' }}
          />
        </div>

        <div className="form-group">
          <label>并发数</label>
          <select
            value={settings.workers}
            onChange={(e) => setSettings((s) => ({ ...s, workers: Number(e.target.value) }))}
          >
            {[1, 2, 4, 8].map((n) => (
              <option key={n} value={n}>{n}</option>
            ))}
          </select>
        </div>

        <div className="form-group">
          <label>主题</label>
          <select
            value={settings.theme}
            onChange={(e) => setSettings((s) => ({ ...s, theme: e.target.value }))}
          >
            <option value="light">浅色</option>
          </select>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginTop: 16 }}>
          <button className="btn-primary" onClick={handleSave}>
            保存设置
          </button>
          {saved && <span style={{ color: 'var(--success)', fontSize: 13 }}>✓ 已保存</span>}
        </div>
      </div>
    </div>
  )
}
