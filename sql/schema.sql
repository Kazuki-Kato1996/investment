-- 銘柄マスタ
CREATE TABLE IF NOT EXISTS stocks (
    code        VARCHAR(10) PRIMARY KEY,
    name        VARCHAR(100) NOT NULL,
    sector      VARCHAR(50),
    fiscal_month INTEGER,          -- 決算月（3, 6, 9, 12等）
    currency    CHAR(3) DEFAULT 'JPY',  -- JPY or USD
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 売買記録（購入価額シート）
CREATE TABLE IF NOT EXISTS trades (
    id          SERIAL PRIMARY KEY,
    code        VARCHAR(10) REFERENCES stocks(code),
    trade_date  DATE NOT NULL,
    trade_type  VARCHAR(4) NOT NULL CHECK (trade_type IN ('BUY', 'SELL')),
    price       NUMERIC(12, 2) NOT NULL,
    quantity    INTEGER NOT NULL,
    account     VARCHAR(50),       -- 岡地現金, 住信SBI, SBI米株 等
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (code, trade_date, trade_type, price, quantity, account)
);

-- 月次スナップショット（2026.x シート）
CREATE TABLE IF NOT EXISTS monthly_snapshots (
    id              SERIAL PRIMARY KEY,
    code            VARCHAR(10) REFERENCES stocks(code),
    snapshot_month  DATE NOT NULL,             -- その月の1日（例：2026-03-01）
    current_price   NUMERIC(12, 2),
    market_value    NUMERIC(14, 2),            -- 評価額
    unrealized_gain NUMERIC(14, 2),            -- 含み損益
    dividend        NUMERIC(10, 2),            -- 配当金
    account         VARCHAR(50),
    UNIQUE (code, snapshot_month, account)
);

-- 年間集計（年間集計シート）
CREATE TABLE IF NOT EXISTS annual_summaries (
    id              SERIAL PRIMARY KEY,
    account         VARCHAR(50) NOT NULL,
    year            INTEGER NOT NULL,
    month           INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    cash            NUMERIC(14, 2),
    evaluation_gain NUMERIC(14, 2),            -- 評価益
    realized_gain   NUMERIC(14, 2),            -- 売却損益
    total_assets    NUMERIC(14, 2),            -- 資産合計
    UNIQUE (account, year, month)
);

-- 市場指標（業種PER・空売り比率・信用倍率等）
CREATE TABLE IF NOT EXISTS market_indicators (
    id            SERIAL PRIMARY KEY,
    recorded_date DATE NOT NULL,
    sector        VARCHAR(50) NOT NULL,
    per           NUMERIC(8, 2),
    short_ratio   NUMERIC(6, 2),              -- 空売り比率（%）
    margin_ratio  NUMERIC(8, 2),              -- 信用倍率
    UNIQUE (recorded_date, sector)
);

-- 株価履歴（fetch_prices.py で定期取得）
CREATE TABLE IF NOT EXISTS price_history (
    id          SERIAL PRIMARY KEY,
    code        VARCHAR(10) REFERENCES stocks(code),
    price_date  DATE NOT NULL,
    open_price  NUMERIC(12, 2),
    high_price  NUMERIC(12, 2),
    low_price   NUMERIC(12, 2),
    close_price NUMERIC(12, 2),
    volume      BIGINT,
    UNIQUE (code, price_date)
);

-- updated_at 自動更新トリガー
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE TRIGGER stocks_updated_at
    BEFORE UPDATE ON stocks
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();
