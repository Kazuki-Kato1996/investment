"""
売買表2026.xlsx → PostgreSQL マイグレーションスクリプト

使い方:
    pip install -r requirements.txt
    docker compose up -d
    python scripts/migrate.py
"""

import os
import re
from datetime import datetime, date

import openpyxl
import psycopg2
from psycopg2.extras import execute_values

XLSX_PATH = os.path.join(os.path.dirname(__file__), '..', '売買表2026.xlsx')

DB_CONFIG = {
    'host': 'localhost',
    'port': 5432,
    'dbname': 'investment_db',
    'user': 'investor',
    'password': 'investment2026',
}

# 口座名マッピング（年間集計シートの列→口座名）
ACCOUNT_COLUMNS = {
    'B': '岡地現金',
    'C': '住信SBI',
    'D': 'SBI米株',
}


def load_workbook(path: str):
    print(f"Excelファイルを読み込み中: {path}")
    return openpyxl.load_workbook(path, data_only=True)


def cell_value(ws, row: int, col: str):
    """列アルファベット + 行番号でセル値を取得"""
    return ws[f'{col}{row}'].value


def to_date(value) -> date | None:
    """Excelの日付値またはdatetimeをdateに変換"""
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value if isinstance(value, date) else value.date()
    if isinstance(value, (int, float)):
        # Excelシリアル値
        from openpyxl.utils.datetime import from_excel
        return from_excel(value).date()
    return None


def to_float(value) -> float | None:
    try:
        return float(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def to_int(value) -> int | None:
    try:
        return int(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def migrate_stocks(conn, wb):
    """銘柄マスタを購入価額シートから移行"""
    print("銘柄マスタを移行中...")
    ws = wb['購入価額']
    stocks = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        code = str(row[0]).strip() if row[0] else None   # B列: コード
        name = str(row[1]).strip() if row[1] else None   # C列: 銘柄名
        if not code or not name or not re.match(r'^\d+$', code):
            continue
        stocks[code] = {'code': code, 'name': name, 'currency': 'JPY'}

    # SBI米株は USD
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = str(row[0]).strip() if row[0] else None
        account = str(row[5]).strip() if len(row) > 5 and row[5] else ''
        if code and 'SBI米株' in account:
            if code in stocks:
                stocks[code]['currency'] = 'USD'

    with conn.cursor() as cur:
        execute_values(
            cur,
            """
            INSERT INTO stocks (code, name, currency)
            VALUES %s
            ON CONFLICT (code) DO UPDATE SET name = EXCLUDED.name
            """,
            [(s['code'], s['name'], s['currency']) for s in stocks.values()]
        )
    conn.commit()
    print(f"  → {len(stocks)} 銘柄を登録")


def migrate_trades(conn, wb):
    """売買記録を購入価額シートから移行"""
    print("売買記録を移行中...")
    ws = wb['購入価額']
    trades = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        code = str(row[0]).strip() if row[0] else None
        if not code or not re.match(r'^\d+$', code):
            continue
        trade_date = to_date(row[2])   # D列: 日付
        price = to_float(row[3])       # E列: 株価
        quantity = to_int(row[4]) if len(row) > 4 else None
        account = str(row[5]).strip() if len(row) > 5 and row[5] else None

        if trade_date and price and quantity:
            trades.append((code, trade_date, 'BUY', price, quantity, account))

    with conn.cursor() as cur:
        execute_values(
            cur,
            """
            INSERT INTO trades (code, trade_date, trade_type, price, quantity, account)
            VALUES %s
            ON CONFLICT DO NOTHING
            """,
            trades
        )
    conn.commit()
    print(f"  → {len(trades)} 件の売買記録を登録")


def migrate_monthly_snapshots(conn, wb):
    """月次スナップショットを 2026.x シートから移行"""
    print("月次スナップショットを移行中...")
    total = 0

    for sheet_name in ['2026.1', '2026.2', '2026.3']:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        year, month = map(int, sheet_name.split('.'))
        snapshot_month = date(year, month, 1)
        snapshots = []

        for row in ws.iter_rows(min_row=3, values_only=True):
            code = str(row[0]).strip() if row[0] else None
            if not code or not re.match(r'^\d+$', code):
                continue
            current_price = to_float(row[3])
            market_value = to_float(row[4])
            unrealized_gain = to_float(row[5])
            dividend = to_float(row[6]) if len(row) > 6 else None
            account = str(row[1]).strip() if row[1] else None

            if current_price is not None:
                snapshots.append((code, snapshot_month, current_price,
                                  market_value, unrealized_gain, dividend, account))

        with conn.cursor() as cur:
            execute_values(
                cur,
                """
                INSERT INTO monthly_snapshots
                    (code, snapshot_month, current_price, market_value, unrealized_gain, dividend, account)
                VALUES %s
                ON CONFLICT (code, snapshot_month, account) DO UPDATE SET
                    current_price   = EXCLUDED.current_price,
                    market_value    = EXCLUDED.market_value,
                    unrealized_gain = EXCLUDED.unrealized_gain,
                    dividend        = EXCLUDED.dividend
                """,
                snapshots
            )
        conn.commit()
        print(f"  → {sheet_name}: {len(snapshots)} 件")
        total += len(snapshots)

    print(f"  → 合計 {total} 件")


def migrate_annual_summaries(conn, wb):
    """年間集計を移行"""
    print("年間集計を移行中...")
    ws = wb['年間集計']
    summaries = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        year = to_int(row[0])
        month = to_int(row[1])
        if not year or not month:
            continue
        for col_idx, account in [(2, '岡地現金'), (3, '住信SBI'), (4, 'SBI米株')]:
            total_assets = to_float(row[col_idx]) if len(row) > col_idx else None
            if total_assets is not None:
                summaries.append((account, year, month, None, None, None, total_assets))

    with conn.cursor() as cur:
        execute_values(
            cur,
            """
            INSERT INTO annual_summaries
                (account, year, month, cash, evaluation_gain, realized_gain, total_assets)
            VALUES %s
            ON CONFLICT (account, year, month) DO UPDATE SET
                total_assets = EXCLUDED.total_assets
            """,
            summaries
        )
    conn.commit()
    print(f"  → {len(summaries)} 件の年間集計を登録")


def main():
    wb = load_workbook(XLSX_PATH)
    print(f"シート一覧: {wb.sheetnames}\n")

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        migrate_stocks(conn, wb)
        migrate_trades(conn, wb)
        migrate_monthly_snapshots(conn, wb)
        migrate_annual_summaries(conn, wb)
        print("\n移行完了!")
    finally:
        conn.close()


if __name__ == '__main__':
    main()
