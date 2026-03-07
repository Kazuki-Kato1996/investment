import { useState } from 'react'
import Dashboard from './pages/Dashboard.jsx'
import Stocks from './pages/Stocks.jsx'
import Trades from './pages/Trades.jsx'
import Monthly from './pages/Monthly.jsx'

const PAGES = [
  { key: 'dashboard', label: 'ダッシュボード' },
  { key: 'stocks',    label: '銘柄一覧' },
  { key: 'trades',    label: '売買記録' },
  { key: 'monthly',   label: '月次損益' },
]

export default function App() {
  const [page, setPage] = useState('dashboard')

  return (
    <>
      <nav>
        <h1>投資管理</h1>
        {PAGES.map(p => (
          <button
            key={p.key}
            className={page === p.key ? 'active' : ''}
            onClick={() => setPage(p.key)}
          >
            {p.label}
          </button>
        ))}
      </nav>
      <main>
        {page === 'dashboard' && <Dashboard />}
        {page === 'stocks'    && <Stocks />}
        {page === 'trades'    && <Trades />}
        {page === 'monthly'   && <Monthly />}
      </main>
    </>
  )
}
