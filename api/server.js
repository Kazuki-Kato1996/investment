const express = require('express');
const cors = require('cors');
const { Pool } = require('pg');

const app = express();
app.use(cors());
app.use(express.json());

const pool = new Pool({
  host:     process.env.DB_HOST     || 'localhost',
  port:     process.env.DB_PORT     || 5432,
  database: process.env.DB_NAME     || 'investment_db',
  user:     process.env.DB_USER     || 'investor',
  password: process.env.DB_PASSWORD || 'investment2026',
});

// ダッシュボード: 総資産・含み損益サマリー
app.get('/api/dashboard', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT
      snapshot_month,
      SUM(market_value)    AS total_market_value,
      SUM(unrealized_gain) AS total_unrealized_gain,
      COUNT(DISTINCT code) AS stock_count
    FROM monthly_snapshots
    GROUP BY snapshot_month
    ORDER BY snapshot_month DESC
    LIMIT 12
  `);
  res.json(rows);
});

// 銘柄一覧 + 最新株価
app.get('/api/stocks', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT
      s.code,
      s.name,
      s.sector,
      s.currency,
      ph.close_price AS latest_price,
      ph.price_date  AS price_date
    FROM stocks s
    LEFT JOIN LATERAL (
      SELECT close_price, price_date
      FROM price_history
      WHERE code = s.code
      ORDER BY price_date DESC
      LIMIT 1
    ) ph ON true
    ORDER BY s.code
  `);
  res.json(rows);
});

// 売買記録
app.get('/api/trades', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT
      t.id,
      t.code,
      s.name,
      t.trade_date,
      t.trade_type,
      t.price,
      t.quantity,
      t.price * t.quantity AS total_amount,
      t.account
    FROM trades t
    JOIN stocks s ON t.code = s.code
    ORDER BY t.trade_date DESC
  `);
  res.json(rows);
});

// 月次スナップショット
app.get('/api/monthly', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT
      ms.snapshot_month,
      ms.code,
      s.name,
      ms.account,
      ms.current_price,
      ms.market_value,
      ms.unrealized_gain
    FROM monthly_snapshots ms
    JOIN stocks s ON ms.code = s.code
    ORDER BY ms.snapshot_month DESC, ms.unrealized_gain DESC
  `);
  res.json(rows);
});

// 年間集計
app.get('/api/annual', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT account, year, month, total_assets, evaluation_gain
    FROM annual_summaries
    ORDER BY year DESC, month DESC
  `);
  res.json(rows);
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`API server running on http://localhost:${PORT}`));
