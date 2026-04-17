# WorkingTimeChecker 自動処理ツール

## 概要

本ツールは、以下の一連の月次業務を自動で実行するための PowerShell + Python ベースの業務自動化ツールです。

- Confluence からの添付ファイルダウンロード
- Excel マクロによる工数集計処理
- Apps(Web勤怠) からの勤怠データ取得
- 工数表と Web 勤怠の不整合チェック
- Confluence ページの更新・コピー
- Rocket.Chat への完了通知

**月次（YYYY_MM）切替にも対応しており、設定変更なしで継続運用可能です。**

---

## ディレクトリ構成
```
WorkingTimeChecker/
│
├─ run.ps1                   # 実行用エントリポイント（本 README の対象）
├─ Run-ExcelMacro.ps1        # Excel VBA マクロ実行専用
│
├─ config/
│   ├─ .env                  # 秘密情報（Git 管理しない）
│   ├─ .env.example          # .env のテンプレート
│   └─ config.yaml           # 各種設定（Git 管理対象）
│
├─ ConfluenceDownload.py
├─ WebAttendanceDownload.py
├─ WorkingTimeChecker.py
├─ CopyConfluence.py
├─ PostRocketChatMessage.py
│
├─ config_loader.py
├─ month_utils.py
├─ path_utils.py
└─ confluence_utils.py
```

---

## 事前準備

### 1. Python 環境

- Python 3.9 以上を推奨
- 必要ライブラリは事前にインストールしておく

---
## requirements.txt について

本ツールで使用している Python ライブラリは `requirements.txt` にまとめています。  
各ライブラリの用途は以下のとおりです。

### Core / 設定関連

- **PyYAML**  
  - `config.yaml` の読み込みに使用  
  - `config_loader.py` で利用

---

### Confluence 連携

- **atlassian-python-api**  
  - Confluence のページ取得、更新、添付操作  
  - 使用箇所：
    - `ConfluenceDownload.py`
    - `CopyConfluence.py`
    - `WorkingTimeChecker.py`

---

### Rocket.Chat 連携

- **rocketchat-API**  
  - Rocket.Chat へのメッセージ投稿  
  - 使用箇所：
    - `PostRocketChatMessage.py`

---

### データ処理

- **pandas**  
  - CSV / Excel データの読み込み・集計  
  - 工数表と Web 勤怠の不整合チェック  
  - HTML（Confluence 用）への変換

---

### Excel 取扱い

- **openpyxl**  
  - Excel（`.xlsm`）ファイルの読み書き  
  - 主に集計結果の保存に使用

- **pywin32（Windows 環境のみ）**  
  - Excel COM 操作・Outlook メール送信  
  - Windows 環境限定のため、条件付きで指定

---

### Web 操作（Apps / Web勤怠）

- **selenium**  
  - Web 勤怠画面の自動操作  
  - CSV ダウンロード処理に使用

- **webdriver-manager**  
  - ChromeDriver の自動取得・管理  
  - 環境差異によるドライバ不整合を回避

---

### 通信系（補助）

- **requests / urllib3**  
  - API 通信の内部依存ライブラリ  
  - 明示指定することでバージョン差異による事故を防止

---

## requirements.txt のインストール方法

```powershell
pip install -r requirements.txt
```
- Python 実行前に一度だけ実施してください
- Chrome がインストールされている必要があります（Web 勤怠取得用）

---
### 2. .env の設定

config/.env.example をコピーして .env を作成し、以下の情報を設定してください。
```
DNDEV_AWS_ID=xxx
DNDEV_AWS_PASSWORD=xxx
GENIIE_ID=xxx
GENIIE_PASSWORD=xxx

APPS_ID=xxx
APPS_PASSWORD=xxx
APPS_TOTP_SECRET=xxx

ROCKETCHAT_USER_ID=xxx
ROCKETCHAT_TOKEN=xxx
```
⚠ .env は **Git 管理対象外**です。

---
### 3. config.yaml の設定

年月の切替やパス・Excel マクロ名の定義は、すべて config.yaml で行います。

### 月次指定

```yaml
system:
  target_month: auto    # auto / YYYY-MM 形式（例: 2026-04）
```
- auto：実行日基準の年月を自動使用
- YYYY-MM：過去月・特定月の再実行用

### Excel マクロ定義

```yaml
excel:
  macros:
    - "工数管理表読み込み"
```
- 上から順に実行を試み、成功したものを採用
- マクロは Public / 引数なし で定義されている必要があります

---

## 実行方法

### ▶ 通常実行（すべての処理を順番に実行）

```powershell
.\run.ps1
```

---
### 部分実行・再実行
本ツールは **再実行・途中実行を前提**に設計されています。

### ▶ 指定ステップ「以降」を実行（途中再開）

```powershell
# Excel マクロ以降を再実行
.\run.ps1 -From ExcelMacro

# WorkingTimeChecker 以降を再実行
.\run.ps1 -From WorkingTimeChecker
```

### ▶ 指定ステップ「のみ」を実行

```powershell
# Excel マクロのみ
.\run.ps1 -Only ExcelMacro

# Confluence コピーのみ
.\run.ps1 -Only CopyConfluence
```

### ▶ 開始〜終了ステップを指定

```powershell
# WebAttendanceDownload ～ CopyConfluence まで
.\run.ps1 -From WebAttendanceDownload -To CopyConfluence
```

---
## ステップ一覧（実行順）

|ステップ名|内容|
|:-----------|:------------:|
|ConfluenceDownload|Confluence から添付ファイルを取得|
|ExcelMacro|Excel VBA マクロによる集計処理|
|WebAttendanceDownload|Apps(Web勤怠) から CSV 取得|
|WorkingTimeChecker|工数表と Web 勤怠の不整合チェック|
|CopyConfluence|Confluence ページのコピー|
|PostRocketChat|Rocket.Chat へ完了通知|

---
## エラー時の挙動

- いずれかのステップでエラーが発生した場合、その時点で処理を中断します
- エラー内容は PowerShell の標準エラー出力に表示されます
- 必要に応じて -From オプションで **該当ステップから再実行**してください

---
## 注意事項

- Excel マクロ処理は時間がかかる場合があります（数分〜十数分）
- 実行中は Excel を操作しないでください
- ネットワークパス（UNC）を使用するため、通信環境に依存します

---
## 設計思想

- **run.ps1**：処理の流れと実行順のみを管理
- **Python**：ロジックと設定解釈を担当
- **config.yaml**：設定を一元管理
- **.env**：秘密情報のみ保持
- **再実行・復旧しやすさを最優先**

---
## 更新履歴

- config.yaml/.envによる設定に対応
- 月次自動化（YYYY_MM 自動生成）対応
- 処理ステップ選択実行対応（From / To / Only）
- Excel マクロ名の外部定義対応
- Rocket.Chat 通知統合