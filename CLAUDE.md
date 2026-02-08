# 売上管理統合システム（GAS）

## プロジェクト概要
Google Apps Script上の売上管理システム。スプレッドシートをDBとし、7段階パイプライン管理・帳票PDF生成・Webダッシュボードを提供。

- **リポジトリ**: https://github.com/K-nagami/gas-sales-management
- **プラットフォーム**: Google Apps Script（コンテナバインドスクリプト）
- **フロントエンド**: webapp.html（単一HTML、Chart.js使用）
- **開発ログ**: CHANGELOG.md を参照

## ファイル構成

| ファイル | 行数 | 役割 |
|---|---|---|
| `setup.gs` | 628行 | 初期構築、migrateV2/V3、サンプルデータ |
| `main.gs` | 811行 | PDF生成エンジン、CSV出力、自動採番、onEditトリガー |
| `templates.gs` | 273行 | 見積書/発注書/請求書のGoogleDocテンプレート生成 |
| `webapp.gs` | ~860行 | Web App API（ダッシュボード集計、CRUD、削除、一括PDF生成） |
| `webapp.html` | ~960行 | ダッシュボードUI（KPI、ファネル、案件管理、新規登録） |
| `DEPLOY_GUIDE.md` | | デプロイ手順書 |

## スプレッドシート構造

### 案件マスタ（A〜T列、20列）
```
A=案件ID(EST-YYYYMM-NNN), B=顧客名, C=案件名, D=見積日, E=受注日,
F=請求日, G=入金予定日, H=入金日, I=ステータス(7段階),
J=小計(SUMIF), K=消費税(FLOOR), L=合計(J+K),
M=見積書URL, N=発注書URL, O=請求書URL, P=備考, Q=ヨミランク(A/B/C/D),
R=登録日, S=提案日, T=失注日
```

### 明細（A〜E列）
```
A=案件ID(FK), B=品目名, C=単価, D=数量, E=小計(=C*D)
```

### 設定（B1〜B19）
```
B1=自社名, B2=住所, ..., B13=税率(0.10), B16=見積テンプレID,
B17=発注テンプレID, B18=請求テンプレID, B19=出力フォルダID
```

## 7段階ステータス
商談見込み → 提案中 → 見積もり提示 → 受注 → 請求済 → 入金済 / 失注

## 会計年度
決算月=1月、期首=2月（2月〜翌1月が1期）

## v3 基準日方式
KPI集計は「現ステータスの基準日」で期間フィルタする。
- 商談見込み→R列(登録日), 提案中→S列(提案日), 見積もり提示→D列(見積日)
- 受注→E列(受注日), 請求済→F列(請求日), 入金済→H列(入金日), 失注→T列(失注日)

## 主要関数（バックエンド）

### webapp.gs
- `doGet()` — Web Appエントリー
- `getDashboardData(period)` — KPI/ファネル/月別/入金ギャップ集計
- `getProjectsList()` / `getItemsForProjectWeb(projectId)` — 一覧・明細取得
- `addNewProject(formData)` / `addNewItem(formData)` — 新規登録
- `updateProjectBulk(projectId, updates)` — ステータス/ヨミ/備考一括更新
- `generatePDFFromWeb(projectId, docType)` — 単体PDF生成
- `generateBulkInvoices(projectIds)` — 一括請求書PDF生成
- `deleteProject(projectId)` — 案件+明細削除

### main.gs
- `fillTemplateAndConvertToPDF(templateId, replacements, items, fileName, folderId)` — PDF生成コア
- `buildReplacements_(rowData, settings, items)` — プレースホルダ辞書
- `getItemsForProject(projectId)` — 明細取得
- `getSettingsData()` — 設定シート読込
- `exportCSVForAccountant()` — 税理士CSV出力
- `onEdit(e)` — ステータス変更時の日付自動入力

## コーディング規約
- GAS（ES5ベース、var使用、letやconstは使わない）
- `google.script.run.withSuccessHandler().withFailureHandler()` でフロントと通信
- 戻り値パターン: `{ success: true/false, message: '...', ... }`
- プライベート関数は末尾に `_` を付ける（例: `formatDate_`）
- スプレッドシート操作: getDataRange().getValues() で一括取得、個別セルはgetRange().setValue()
- 行削除は下→上（インデックスずれ防止）

## デプロイ手順
1. GASエディタで全ファイルを貼り付け
2. `setup()` を実行（スプレッドシート初期構築）
3. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」→アクセス=自分のみ
4. 既存環境: `migrateV3()` を先に実行

## GitHub操作
```bash
git remote: https://github.com/K-nagami/gas-sales-management.git
ブランチ: main
```
