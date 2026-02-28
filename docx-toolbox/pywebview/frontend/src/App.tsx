import { Routes, Route, Navigate } from 'react-router-dom'
import Layout from './components/Layout'
import Dashboard from './pages/Dashboard'
import ExcelToolPage from './pages/ExcelToolPage'
import ImageToolPage from './pages/ImageToolPage'
import TableToolPage from './pages/TableToolPage'
import BatchPage from './pages/BatchPage'
import SettingsPage from './pages/SettingsPage'

export default function App() {
  return (
    <Routes>
      <Route element={<Layout />}>
        <Route index element={<Dashboard />} />
        <Route path="excel" element={<ExcelToolPage />} />
        <Route path="image" element={<ImageToolPage />} />
        <Route path="table" element={<TableToolPage />} />
        <Route path="batch" element={<BatchPage />} />
        <Route path="settings" element={<SettingsPage />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Route>
    </Routes>
  )
}
