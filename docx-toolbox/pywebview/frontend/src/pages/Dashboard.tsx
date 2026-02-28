import { useEffect, useState } from 'react'
import { Link } from 'react-router-dom'
import { api } from '../api/bridge'
import type { TaskInfo } from '../types'

export default function Dashboard() {
  const [tasks, setTasks] = useState<TaskInfo[]>([])

  useEffect(() => {
    const bridge = api()
    if (!bridge) return
    bridge.list_tasks().then((res) => {
      if (res.ok && res.data) setTasks(res.data)
    })
  }, [])

  return (
    <div>
      <h1 className="page-title">仪表盘</h1>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16, marginBottom: 24 }}>
        <Link to="/excel" style={{ textDecoration: 'none' }}>
          <div className="card" style={{ textAlign: 'center' }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>📑</div>
            <div style={{ fontWeight: 600 }}>Excel 处理</div>
            <div style={{ fontSize: 13, color: 'var(--text-secondary)', marginTop: 4 }}>
              嵌入对象 All-in-One
            </div>
          </div>
        </Link>
        <Link to="/image" style={{ textDecoration: 'none' }}>
          <div className="card" style={{ textAlign: 'center' }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>🖼️</div>
            <div style={{ fontWeight: 600 }}>图片分离</div>
            <div style={{ fontSize: 13, color: 'var(--text-secondary)', marginTop: 4 }}>
              分离并标记文档图片
            </div>
          </div>
        </Link>
        <Link to="/table" style={{ textDecoration: 'none' }}>
          <div className="card" style={{ textAlign: 'center' }}>
            <div style={{ fontSize: 32, marginBottom: 8 }}>📋</div>
            <div style={{ fontWeight: 600 }}>表格提取</div>
            <div style={{ fontSize: 13, color: 'var(--text-secondary)', marginTop: 4 }}>
              提取并导出文档表格
            </div>
          </div>
        </Link>
      </div>

      <h2 style={{ fontSize: 16, fontWeight: 600, marginBottom: 12 }}>最近任务</h2>
      {tasks.length === 0 ? (
        <div className="card" style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: 32 }}>
          暂无任务记录
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {tasks.slice(0, 10).map((task) => (
            <div key={task.task_id} className="card" style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px' }}>
              <span className={`status-badge ${task.status}`}>{task.status}</span>
              <span style={{ fontWeight: 500, fontSize: 14 }}>{task.task_type}</span>
              <span style={{ color: 'var(--text-secondary)', fontSize: 13, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {task.input_path}
              </span>
              <span style={{ color: 'var(--text-secondary)', fontSize: 12 }}>
                {task.created_at}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
