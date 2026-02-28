import { NavLink, Outlet } from 'react-router-dom'
import LogPanel from './LogPanel'

const navItems = [
  { to: '/', label: '仪表盘', icon: '📊' },
  { to: '/excel', label: 'Excel 处理', icon: '📑' },
  { to: '/image', label: '图片分离', icon: '🖼️' },
  { to: '/table', label: '表格提取', icon: '📋' },
  { to: '/batch', label: '批处理', icon: '⚡' },
  { to: '/settings', label: '设置', icon: '⚙️' },
]

export default function Layout() {
  return (
    <div style={{ display: 'flex', height: '100vh', flexDirection: 'column' }}>
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        {/* Sidebar */}
        <nav style={{
          width: 'var(--sidebar-width)',
          minWidth: 'var(--sidebar-width)',
          background: 'var(--bg-sidebar)',
          borderRight: '1px solid var(--border)',
          display: 'flex',
          flexDirection: 'column',
          padding: '16px 0',
        }}>
          <div style={{
            padding: '0 16px 20px',
            fontSize: 16,
            fontWeight: 700,
            color: 'var(--primary)',
          }}>
            DOCX 工具箱
          </div>
          {navItems.map((item) => (
            <NavLink
              key={item.to}
              to={item.to}
              end={item.to === '/'}
              style={({ isActive }) => ({
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                padding: '10px 16px',
                fontSize: 14,
                color: isActive ? 'var(--primary)' : 'var(--text-primary)',
                background: isActive ? 'rgba(59,130,246,0.08)' : 'transparent',
                borderRight: isActive ? '3px solid var(--primary)' : '3px solid transparent',
                textDecoration: 'none',
                transition: 'background 0.15s',
              })}
            >
              <span>{item.icon}</span>
              <span>{item.label}</span>
            </NavLink>
          ))}
        </nav>

        {/* Main content */}
        <main style={{
          flex: 1,
          overflow: 'auto',
          padding: 24,
        }}>
          <Outlet />
        </main>
      </div>

      {/* Log panel */}
      <LogPanel />
    </div>
  )
}
