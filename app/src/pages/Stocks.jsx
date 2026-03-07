import { useEffect, useState } from 'react'

export default function Stocks() {
  const [stocks, setStocks] = useState([])
  const [loading, setLoading] = useState(true)
  const [search, setSearch] = useState('')

  useEffect(() => {
    fetch('/api/stocks')
      .then(r => r.json())
      .then(d => { setStocks(d); setLoading(false) })
  }, [])

  if (loading) return <div className="loading">読み込み中...</div>

  const filtered = stocks.filter(s =>
    s.code.includes(search) || s.name?.includes(search)
  )

  return (
    <>
      <h2>銘柄一覧</h2>
      <input
        type="text"
        placeholder="コード・銘柄名で検索"
        value={search}
        onChange={e => setSearch(e.target.value)}
        style={{ marginBottom: 16, padding: '8px 12px', borderRadius: 6, border: '1px solid #ddd', width: 260, fontSize: 14 }}
      />
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>コード</th>
              <th>銘柄名</th>
              <th>通貨</th>
              <th>最新株価</th>
              <th>取得日</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map(s => (
              <tr key={s.code}>
                <td>{s.code}</td>
                <td>{s.name}</td>
                <td>{s.currency}</td>
                <td>{s.latest_price != null ? Number(s.latest_price).toLocaleString('ja-JP') : '—'}</td>
                <td>{s.price_date ? new Date(s.price_date).toLocaleDateString('ja-JP') : '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  )
}
