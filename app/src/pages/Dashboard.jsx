import { useEffect, useState } from 'react'

function fmt(n) {
  if (n == null) return '—'
  return Number(n).toLocaleString('ja-JP') + ' 円'
}

export default function Dashboard() {
  const [data, setData] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    fetch('/api/dashboard')
      .then(r => r.json())
      .then(d => { setData(d); setLoading(false) })
  }, [])

  if (loading) return <div className="loading">読み込み中...</div>

  const latest = data[0] || {}

  return (
    <>
      <h2>ダッシュボード</h2>

      <div className="cards">
        <div className="card">
          <div className="label">時価評価額（直近月）</div>
          <div className="value">{fmt(latest.total_market_value)}</div>
        </div>
        <div className="card">
          <div className="label">含み損益（直近月）</div>
          <div className={`value ${Number(latest.total_unrealized_gain) >= 0 ? 'positive' : 'negative'}`}>
            {fmt(latest.total_unrealized_gain)}
          </div>
        </div>
        <div className="card">
          <div className="label">保有銘柄数</div>
          <div className="value">{latest.stock_count ?? '—'} 銘柄</div>
        </div>
        <div className="card">
          <div className="label">直近月</div>
          <div className="value" style={{ fontSize: 16 }}>
            {latest.snapshot_month ? new Date(latest.snapshot_month).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long' }) : '—'}
          </div>
        </div>
      </div>

      <h2 style={{ marginTop: 8 }}>月次推移</h2>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>月</th>
              <th>時価評価額</th>
              <th>含み損益</th>
              <th>保有銘柄数</th>
            </tr>
          </thead>
          <tbody>
            {data.map(row => (
              <tr key={row.snapshot_month}>
                <td>{new Date(row.snapshot_month).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long' })}</td>
                <td>{fmt(row.total_market_value)}</td>
                <td className={Number(row.total_unrealized_gain) >= 0 ? 'positive' : 'negative'}>
                  {fmt(row.total_unrealized_gain)}
                </td>
                <td>{row.stock_count} 銘柄</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  )
}
