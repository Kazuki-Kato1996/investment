# Investment Project

## 概要
株式投資に関する情報収集・分析を行うプロジェクト。

## 目的
- 株式市場のデータ収集
- 投資判断に役立つ分析の実施
- 投資関連情報の整理・可視化

## 技術スタック
- Node.js / JavaScript（メイン）
- Python（データ収集・分析スクリプト）

## ディレクトリ構成
```
investment/
├── api/                  # APIサーバー（Express）
├── app/                  # フロントエンド
├── morning-report/       # 朝の市場レポート自動送信
├── scripts/              # Pythonスクリプト（価格取得・マイグレーション）
├── spreadsheet-updater/  # Googleスプレッドシート更新・レポート生成
└── sql/                  # DBスキーマ
```

## コーディング規約
- コードのコメントは日本語
- JavaScript: ESM（import/export）を使用
- Python: スネークケース（snake_case）
