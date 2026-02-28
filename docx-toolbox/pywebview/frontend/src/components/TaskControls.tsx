interface TaskControlsProps {
  running: boolean
  onStart: () => void
  onCancel: () => void
  disabled?: boolean
}

export default function TaskControls({ running, onStart, onCancel, disabled }: TaskControlsProps) {
  return (
    <div style={{ display: 'flex', gap: 10, marginTop: 16 }}>
      {running ? (
        <button className="btn-danger" onClick={onCancel}>
          取消任务
        </button>
      ) : (
        <button className="btn-primary" onClick={onStart} disabled={disabled}>
          开始处理
        </button>
      )}
    </div>
  )
}
