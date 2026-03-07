import { useEffect, useState } from 'react'

export default function Trades() {
  const [trades, setTrades] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    fetch('/api/trades')
      .then(r => r.json())
      .then(d => { setTrades(d); setLoading(false) })
  }, [])

  if (loading) return <div className="loading">読み込み中...</div>

  return (
    <>
      <h2>売買記録</h2>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>日付</th>
              <th>コード</th>
              <th>銘柄名</th>
              <th>種別</th>
              <th>株価</th>
              <th>株数</th>
              <th>約定金額</th>
              <th>口座</th>
            </tr>
          </thead>
          <tbody>
            {trades.map(t => (
              <tr key={t.id}>
                <td>{t.trade_date ? new Date(t.trade_date).toLocaleDateString('ja-JP') : '—'}</td>
                <td>{t.code}</td>
                <td>{t.name}</td>
                <td>
                  <span className={`badge ${t.trade_type === 'BUY' ? 'buy' : 'sell'}`}>
                    {t.trade_type === 'BUY' ? '買い' : '売り'}
                  </span>
                </td>
                <td>{Number(t.price).toLocaleString('ja-JP')}</td>
                <td>{Number(t.quantity).toLocaleString('ja-JP')}</td>
                <td>{Number(t.total_amount).toLocaleString('ja-JP')} 円</td>
                <td>{t.account ?? '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  )
}
