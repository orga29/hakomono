# TROUBLESHOOTING

このファイルでは、開発中や実行時によく遭遇する問題と簡単な対処をまとめます。

## 1) 起動時に "No module named mcp_server_time" と出る
- 原因: 実行設定が `mcp_server_time` モジュールを参照しているが、そのファイルが存在しない。
- 対処:
  - 実行したいファイルが `app.py` などであれば、`streamlit run app.py` や `python app.py` で起動する。
  - 必要なモジュールが別リポジトリにある場合はプロジェクトへ追加するか、`pip install <package>` でインストールする。

## 2) "Port 8501 is already in use"
- 原因: 既に同じポートでプロセスが待ち受けている。
- 対処:
  - 使用中の PID を確認: `netstat -ano | Select-String ':8501'`（PowerShell）
  - プロセス停止（停止して良い場合）: `Stop-Process -Id <PID> -Force`
  - 一時回避: 別ポートで起動: `streamlit run app.py --server.port 8502`

## 3) 依存パッケージが足りない / import error
- 対処:
  - 仮想環境を作成して `pip install -r requirements.txt` を実行する。
  - 特定のパッケージ名がエラーになっている場合は `pip install <package>`。

## 4) 権限エラー（ファイルアクセスやネットワーク）
- 対処:
  - 管理者権限で PowerShell を起動して試す。
  - ファイルのパスや `.env` の読み取り権限を確認する。

## 5) ログの取得
- Streamlit の起動ログをファイルに出したいとき:
```powershell
streamlit run app.py 2>&1 | Tee-Object -FilePath streamlit_start.log
```

---
上記で解決しない場合は、発生しているエラーログをこのリポジトリの Issue に貼ってください。
