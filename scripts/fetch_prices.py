"""
Yahoo Finance API で株価を取得し price_history テーブルへ保存するスクリプト

使い方:
    # 全銘柄の直近1日分を取得
    python scripts/fetch_prices.py

    # 期間を指定して取得
    python scripts/fetch_prices.py --start 2026-01-01 --end 2026-03-07
"""

import argparse
from datetime import date, timedelta

import psycopg2
from psycopg2.extras import execute_values
import yfinance as yf

DB_CONFIG = {
    'host': 'localhost',
    'port': 5432,
    'dbname': 'investment_db',
    'user': 'investor',
    'password': 'investment2026',
}


def get_all_stocks(conn) -> list[dict]:
    """DBから全銘柄を取得"""
    with conn.cursor() as cur:
        cur.execute("SELECT code, currency FROM stocks ORDER BY code")
        return [{'code': row[0], 'currency': row[1]} for row in cur.fetchall()]


def to_ticker(code: str, currency: str) -> str:
    """銘柄コードをYahoo Financeのティッカーに変換"""
    if currency == 'JPY':
        return f"{code}.T"
    return code  # 米株はコードそのまま


def fetch_prices(tickers: list[str], start: date, end: date) -> dict:
    """Yahoo Finance から株価を一括取得"""
    print(f"Yahoo Finance から {len(tickers)} 銘柄を取得中 ({start} 〜 {end})...")
    data = yf.download(
        tickers=tickers,
        start=start.isoformat(),
        end=(end + timedelta(days=1)).isoformat(),
        group_by='ticker',
        auto_adjust=True,
        progress=False,
    )
    return data


def save_prices(conn, records: list[tuple]):
    """price_history テーブルへ UPSERT"""
    with conn.cursor() as cur:
        execute_values(
            cur,
            """
            INSERT INTO price_history (code, price_date, open_price, high_price, low_price, close_price, volume)
            VALUES %s
            ON CONFLICT (code, price_date) DO UPDATE SET
                open_price  = EXCLUDED.open_price,
                high_price  = EXCLUDED.high_price,
                low_price   = EXCLUDED.low_price,
                close_price = EXCLUDED.close_price,
                volume      = EXCLUDED.volume
            """,
            records
        )
    conn.commit()


def main():
    parser = argparse.ArgumentParser(description='Yahoo Finance から株価を取得してDBに保存')
    parser.add_argument('--start', type=date.fromisoformat, default=date.today() - timedelta(days=1))
    parser.add_argument('--end',   type=date.fromisoformat, default=date.today())
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        stocks = get_all_stocks(conn)
        if not stocks:
            print("銘柄マスタが空です。先に migrate.py を実行してください。")
            return

        # ティッカー → 銘柄コードのマッピング
        ticker_to_code = {to_ticker(s['code'], s['currency']): s['code'] for s in stocks}
        tickers = list(ticker_to_code.keys())

        data = fetch_prices(tickers, args.start, args.end)

        records = []
        for ticker, code in ticker_to_code.items():
            try:
                if len(tickers) == 1:
                    df = data
                else:
                    df = data[ticker]
                df = df.dropna(subset=['Close'])
                for idx, row in df.iterrows():
                    records.append((
                        code,
                        idx.date(),
                        float(row['Open'])   if 'Open'   in row else None,
                        float(row['High'])   if 'High'   in row else None,
                        float(row['Low'])    if 'Low'    in row else None,
                        float(row['Close'])  if 'Close'  in row else None,
                        int(row['Volume'])   if 'Volume' in row else None,
                    ))
            except KeyError:
                print(f"  スキップ: {ticker} (データなし)")

        if records:
            save_prices(conn, records)
            print(f"  → {len(records)} 件の株価データを保存しました")
        else:
            print("  取得できたデータがありませんでした")

    finally:
        conn.close()


if __name__ == '__main__':
    main()
