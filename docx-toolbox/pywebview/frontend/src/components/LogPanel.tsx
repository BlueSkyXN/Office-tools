import { useState } from 'react'

interface LogEntry {
  time: string
  level: string
  message: string
}

export default function LogPanel() {
  const [logs] = useState<LogEntry[]>([])
  const [expanded, setExpanded] = useState(false)

  return (
    <div style={{
      height: expanded ? 'var(--log-height)' : 36,
      minHeight: expanded ? 'var(--log-height)' : 36,
      background: 'var(--bg-log)',
      borderTop: '1px solid var(--border)',
      transition: 'height 0.2s',
      display: 'flex',
      flexDirection: 'column',
    }}>
      {/* Header */}
      <div
        onClick={() => setExpanded(!expanded)}
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '6px 16px',
          cursor: 'pointer',
          userSelect: 'none',
          color: 'var(--text-inverse)',
          fontSize: 13,
          fontWeight: 500,
        }}
      >
        <span>📜 日志 ({logs.length})</span>
        <span>{expanded ? '▼' : '▲'}</span>
      </div>

      {/* Log content */}
      {expanded && (
        <div style={{
          flex: 1,
          overflow: 'auto',
          padding: '0 16px 8px',
          fontFamily: 'var(--font-mono)',
          fontSize: 12,
          lineHeight: 1.6,
          color: '#CBD5E1',
        }}>
          {logs.length === 0 ? (
            <div style={{ color: '#64748B', paddingTop: 8 }}>暂无日志</div>
          ) : (
            logs.map((entry, i) => (
              <div key={i}>
                <span style={{ color: '#64748B' }}>{entry.time}</span>{' '}
                <span style={{
                  color: entry.level === 'ERROR' ? 'var(--error)' :
                         entry.level === 'WARN' ? 'var(--warning)' : '#94A3B8',
                }}>[{entry.level}]</span>{' '}
                {entry.message}
              </div>
            ))
          )}
        </div>
      )}
    </div>
  )
}
