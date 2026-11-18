# hellocycling-yokohama-2026

横浜市中区・西区の HELLO CYCLING ポート（25箇所）の台数・空き枠を15分ごとに記録するリポジトリ．

## 自動実行の仕組み

- **起動は外部サービス [cron-job.org](https://cron-job.org/) から行っている．**
  15分ごとに GitHub API（workflow_dispatch）を呼び出して `scrape.yml` を実行している．
  GitHub Actions の `schedule` は数十分〜数時間遅れることがあるため使っていない．
- ワークフローはどちらも `workflow_dispatch` のみなので，**cron-job.org 側が止まると収集も止まる．**

### 止まったときに確認すること

1. cron-job.org のジョブが有効になっているか（失敗が続くと自動で無効化されることがある）
2. cron-job.org に登録している GitHub トークンの期限が切れていないか
   （GitHub → Settings → Developer settings → Personal access tokens）
3. GitHub の Actions タブで `Hello Cycling Scraper` が実行・成功しているか

## ワークフロー

| ファイル | 内容 |
|---|---|
| `.github/workflows/scrape.yml` | データ取得（`scrape.py`）→ `data/` にコミット．日付が変わって最初の回に前日の日次レポートと集計も作る |
| `.github/workflows/summary.yml` | 集計・日次レポートを手動で作り直したいとき用（Actions タブの Run workflow） |

## フォルダ

- `data/` … 15分ごとの生データ（1日1ファイル）
- `reports/` … 日次レポート（Excel）
- `summary/` … 全期間の集計（CSV・Excel）
