# 売上ダッシュボード共有・更新手順

## 1) URLで共有する（GitHub Pages）
1. このフォルダをGitHubリポジトリへpushする。
2. GitHubの `Settings > Pages` で以下を設定する。
   - Source: `Deploy from a branch`
   - Branch: `main`（または公開するブランチ）
   - Folder: `/ (root)`
3. 表示された公開URL（例: `https://<account>.github.io/<repo>/`）を共有する。

## 2) 売上が更新されたときの取り込み
`目標比較` シートを更新したExcelを使って、次のコマンドを実行する。

```powershell
python scripts\generate_data_js.py --xlsx "C:\build\Gpt_project\39期　導管-確報.xlsx" --sheet "目標比較" --out "C:\build\Gpt_project\data.js"
```

このコマンドで `data.js` が最新化される。

## 3) URL反映手順
1. `data.js` を含めてコミット
2. GitHubへpush
3. 1〜2分後に公開URLへ反映

## 補足
- ダッシュボードは `data.js` を優先読み込みします。
- `data.js` がない場合は `app.js` の内蔵データ（フォールバック）を使います。
