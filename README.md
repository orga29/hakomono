# HAKOMONO

小規模な社内用スクリプト集（hakomono）。このリポジトリは主に社内の自動処理や集計に使います。

## 概要
- 目的: 社内でのバッチ処理 / 集計を行う Python スクリプト群
- 主なファイル: `app.py`, `logic.py`, `verify_yamato.py`, `verify_koda.py` など

## 要件
- Python 3.8 以上（推奨: 3.10）
- 依存パッケージは `requirements.txt` を参照

## セットアップ（ローカル環境）
1. 仮想環境を作成して有効化
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```
2. 依存関係をインストール
```powershell
pip install -r requirements.txt
```

## 実行方法（例）
- 単体スクリプトを実行する例:
```powershell
python app.py
```

## データファイル
- Excel（`.xlsm` 等）はリポジトリに含めないことを推奨します。大きなバイナリは別途ファイルサーバか共有ドライブで管理してください。

## 注意点
- 機密情報（APIキーやパスワード）をリポジトリに含めないでください。`.env` を使い、`.gitignore` に追加しています。

## 連絡先
- 問い合わせ: 開発担当 `orga29`（GitHub）または社内チャットで `@○○` に連絡してください。

---
この README は最小限のテンプレートです。必要があれば追記・修正してください。

## クイックスタート
- 仮想環境作成および依存インストール:
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```
- ローカルで起動（開発）:
```powershell
streamlit run app.py
# 別ポートで起動する例
streamlit run app.py --server.port 8502
```

## よく使うコマンド
- 依存の再インストール: `pip install -r requirements.txt`
- ログをファイルに出力して起動:
```powershell
streamlit run app.py 2>&1 | Tee-Object -FilePath streamlit_start.log
```

## トラブルシューティング
- よくある問題と対処は `TROUBLESHOOTING.md` を参照してください。

## 貢献方法
- 開発ルール・PR の流れは `CONTRIBUTING.md` を参照してください。

## 次のステップ（提案）
- 必要なら `LICENSE` を追加します（社内専用表記 or OSS ライセンス）。
- 常駐化（Windows サービスまたはタスクスケジューラ）や CI（GitHub Actions）を追加可能です。
