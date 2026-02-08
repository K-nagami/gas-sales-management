# 売上管理統合システム（GAS）

## What
Google Apps Script製の売上管理システム。スプレッドシートをDB、webapp.htmlでダッシュボード提供。

## Why
- GAS（ES5）: `var`使用、`let`/`const`は使わない
- 単一HTML: webapp.htmlにCSS/JS/HTMLを全て含む（Chart.js CDN使用）
- 会計年度: 決算月=1月、期首=2月（2月〜翌1月が1期）

## Structure
- `gas-project/setup.gs` — 初期構築、マイグレーション
- `gas-project/main.gs` — PDF生成エンジン、CSV出力、onEditトリガー
- `gas-project/templates.gs` — GoogleDocテンプレート生成
- `gas-project/webapp.gs` — Web App API（CRUD、ダッシュボード集計）
- `gas-project/webapp.html` — ダッシュボードUI

## Data Model
- **案件マスタ**: A〜T列（20列）。A=案件ID(EST-YYYYMM-NNN), I=ステータス(7段階), J-L=金額(数式), M-O=帳票URL, Q=ヨミランク, R-T=日付(v3基準日)
- **明細**: A=案件ID(FK), B=品目名, C=単価, D=数量, E=小計(数式)
- **設定**: B1〜B19（自社情報、税率B13、テンプレートID B16-B18、フォルダID B19）

## Rules
- ALWAYS use `{ success: true/false, message: '...' }` as return pattern
- ALWAYS use `google.script.run.withSuccessHandler().withFailureHandler()` for frontend-backend communication
- ALWAYS delete spreadsheet rows bottom-to-top to prevent index shifting
- ALWAYS suffix private functions with `_` (e.g. `formatDate_()`)
- NEVER use `let`, `const`, arrow functions, template literals (ES5 only)
- NEVER modify the 7-stage status order: 商談見込み→提案中→見積もり提示→受注→請求済→入金済/失注

## Deploy
1. GASエディタでgas-project/内の全ファイルを貼り付け
2. `setup()` 実行（初回）/ `migrateV3()` 実行（既存環境）
3. デプロイ → ウェブアプリ → アクセス=自分のみ

## Git
```
repo: https://github.com/K-nagami/gas-sales-management.git
branch: main
```

## References
- `CHANGELOG.md` — 全バージョンの開発ログ
- `DEPLOY_GUIDE.md` — デプロイ詳細手順
