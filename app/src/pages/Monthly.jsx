import { useEffect, useState } from 'react'

function fmt(n) {
  if (n == null) return '—'
  return Number(n).toLocaleString('ja-JP')
}

export default function Monthly() {
  const [data, setData] = useState([])
  const [loading, setLoading] = useState(true)
  const [month, setMonth] = useState('')

  useEffect(() => {
    fetch('/api/monthly')
      .then(r => r.json())
      .then(d => {
        setData(d)
        if (d.length > 0) setMonth(d[0].snapshot_month.slice(0, 7))
        setLoading(false)
      })
  }, [])

  if (loading) return <div className="loading">読み込み中...</div>

  const months = [...new Set(data.map(d => d.snapshot_month.slice(0, 7)))].sort().reverse()
  const filtered = data.filter(d => d.snapshot_month.slice(0, 7) === month)

  const totalGain = filtered.reduce((s, r) => s + (Number(r.unrealized_gain) || 0), 0)
  const totalValue = filtered.reduce((s, r) => s + (Number(r.market_value) || 0), 0)

  return (
    <>
      <h2>月次損益</h2>

      <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 16 }}>
        <select
          value={month}
          onChange={e => setMonth(e.target.value)}
          style={{ padding: '8px 12px', borderRadius: 6, border: '1px solid #ddd', fontSize: 14 }}
        >
          {months.map(m => <option key={m} value={m}>{m}</option>)}
        </select>
        <span style={{ color: '#888', fontSize: 13 }}>{filtered.length} 銘柄</span>
        <span style={{ marginLeft: 'auto', fontWeight: 700 }}>
          評価額合計: <span>{fmt(totalValue)} 円</span>
        </span>
        <span className={totalGain >= 0 ? 'positive' : 'negative'} style={{ fontWeight: 700 }}>
          含み損益: {fmt(totalGain)} 円
        </span>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>コード</th>
              <th>銘柄名</th>
              <th>口座</th>
              <th>現在値</th>
              <th>評価額</th>
              <th>含み損益</th>
            </tr>
          </thead>
          <tbody>
            {filtered.sort((a, b) => Number(b.unrealized_gain) - Number(a.unrealized_gain)).map((r, i) => (
              <tr key={i}>
                <td>{r.code}</td>
                <td>{r.name}</td>
                <td>{r.account ?? '—'}</td>
                <td>{fmt(r.current_price)}</td>
                <td>{fmt(r.market_value)} 円</td>
                <td className={Number(r.unrealized_gain) >= 0 ? 'positive' : 'negative'}>
                  {fmt(r.unrealized_gain)} 円
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  )
}
