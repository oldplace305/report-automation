# 報告書自動化

## 概要
Selenium を使って医療報告書 PDF を自動ダウンロードする Python ツール。
Solamichi (そらみち) サービスから PDF をまとめて取得する。

## 実行方法
```bash
# 事前にChromeをリモートデバッグモードで起動
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 &

# プレビュー（実際にはダウンロードしない）
python3 main.py --dry-run

# 本番実行
python3 main.py
```

## 設定ファイル
`config.json` で対象年月・ダウンロード先を設定する。

## 出力先
`~/Projects/報告書自動化/報告書/`

## 注意
`config.json` の `solamichi_url` は施設固有のURL。外部に公開しないこと。
