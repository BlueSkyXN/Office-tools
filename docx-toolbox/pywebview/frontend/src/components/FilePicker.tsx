import { api } from '../api/bridge'

interface FilePickerProps {
  value: string
  onChange: (path: string) => void
  mode?: 'file' | 'folder'
  label?: string
}

export default function FilePicker({ value, onChange, mode = 'file', label }: FilePickerProps) {
  const handleSelect = async () => {
    const bridge = api()
    if (!bridge) return

    const result = mode === 'folder'
      ? await bridge.select_folder()
      : await bridge.select_input_path()

    if (result.ok && result.data) {
      onChange(result.data)
    }
  }

  return (
    <div className="form-group">
      {label && <label>{label}</label>}
      <div className="form-row">
        <input
          type="text"
          value={value}
          readOnly
          placeholder={mode === 'folder' ? '选择文件夹...' : '选择文件...'}
          style={{ flex: 1 }}
        />
        <button className="btn-secondary" onClick={handleSelect}>
          浏览
        </button>
      </div>
    </div>
  )
}
