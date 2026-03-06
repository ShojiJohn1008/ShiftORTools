# ShiftORTools

病院研修医シフトスケジューリングWebアプリケーション。Google OR-Toolsを用いた制約最適化により、複数病院への研修医配置を自動化します。

## 概要

ShiftORToolsは、日本の病院研修プログラム向けに設計されたシフト管理ツールです。Googleフォームから収集したNG日情報・院外研修スケジュールを読み込み、制約ソルバーで最適なシフト配置を生成します。手動調整・Excelエクスポートにも対応しています。

## 機能

- **病院スロット設定**: 曜日別・日付別の病院受け入れ枠をWeb UIで管理
- **スプレッドシート取り込み**: GoogleフォームのCSV/XLSXを解析し、NG日・院外研修情報を自動抽出
- **自動シフト最適化**: OR-Toolsの制約プログラミングによる最適配置生成（10秒タイムアウト・8並列ワーカー）
- **手動調整**: ドラッグ&ドロップや入力フォームで配置を変更、元に戻す操作にも対応
- **Excelエクスポート**: 病院別・日付別のカラー付きスケジュール表をダウンロード
- **祝日対応**: `jpholiday`ライブラリによる日本の祝日自動判定

## 技術スタック

| 領域 | 技術 |
|------|------|
| バックエンド | Python 3.13+, FastAPI, Uvicorn |
| 最適化エンジン | Google OR-Tools (Constraint Programming) |
| データ処理 | Pandas, OpenPyXL |
| フロントエンド | HTML5, Tailwind CSS, Vanilla JavaScript |
| デプロイ | Heroku (Procfile 同梱) |

## セットアップ

### 前提条件

- Python 3.13 以上
- pip

### インストール

```bash
# リポジトリをクローン
git clone <repository-url>
cd ShiftORTools

# 仮想環境を作成・有効化
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 依存パッケージをインストール
pip install -r requirements.txt
```

### 起動

```bash
# 開発サーバー起動
./scripts/start_dev.sh

# または手動で起動
python -m uvicorn shiftortools.api:app --app-dir src --reload --host 127.0.0.1 --port 8000
```

ブラウザで `http://localhost:8000` を開いてください。

## 使い方

1. **病院設定**: Web UIで各病院の受け入れ枠（曜日別または日付別）を設定
2. **データアップロード**: Googleフォームのエクスポートファイル（Sheet1: NG日、Sheet2: 院外研修）をアップロード
3. **データ確認**: 解析結果・不明な名前・エラーを確認
4. **ソルバー実行**: 「シフト生成」ボタンで最適スケジュールを生成
5. **手動調整**: 必要に応じてドラッグ&ドロップで配置を微調整
6. **エクスポート**: Excelファイルをダウンロードして共有

## 入力スプレッドシート形式

### Sheet1: 研修医NG日

| 列 | 内容 |
|----|------|
| C | 研修医氏名 |
| D | ローテーション種別 |
| G–J | 手動NG日（カンマ/改行/スペース区切り） |

ローテーション種別に応じてNG日が自動生成されます：

| ローテーション | 自動NG |
|---------------|--------|
| 大学病院でローテート | なし |
| 大学外－救急&院外希望 | 平日のみ（祝日除く） |
| 大学外－院外のみ希望 | 全日 |
| 大学外－救急のみ希望 | 平日のみ |
| 大学外－院外も救急も希望しない | 全日 |

### Sheet2: 院外研修スケジュール

| 列 | 内容 |
|----|------|
| A | 日付（空白の場合は直前の行を継承） |
| B | 曜日（参考情報） |
| C | `病院名:氏名` 形式（複数名はカンマ/改行区切り） |

## 配置制約

- 各研修医に割り当てる件数：設定値（デフォルト2件）
- 1日1研修医につき最大1病院
- 大学病院：1研修医あたり最大2回
- 院外病院：1研修医あたり各病院1回まで
- NG日は配置不可

## プロジェクト構成

```
ShiftORTools/
├── frontend/               # Web UI
│   ├── index.html
│   ├── app.js
│   └── styles.css
├── src/shiftortools/       # Pythonバックエンド
│   ├── api.py              # FastAPI エンドポイント
│   ├── solver.py           # OR-Tools 制約ソルバー
│   ├── parsers.py          # CSV/XLSX パーサー
│   ├── schema.py           # データクラス定義
│   ├── utils.py            # ユーティリティ関数
│   └── output.py           # Excel出力
├── scripts/
│   ├── start_dev.sh        # 開発サーバー起動スクリプト
│   └── run_demo.py         # デモスクリプト
├── config/                 # 病院スロット設定（実行時生成）
├── output/                 # ソルバー出力結果（実行時生成）
├── requirements.txt
└── Procfile                # Heroku デプロイ設定
```

## デプロイ（Heroku）

```bash
heroku create your-app-name
git push heroku main
```

## API エンドポイント

| メソッド | パス | 説明 |
|----------|------|------|
| GET | `/api/config` | 病院スロット設定の取得 |
| PUT | `/api/config` | 病院スロット設定の保存 |
| GET | `/api/schedule` | 月次スケジュールのプレビュー |
| GET | `/api/residents` | 解析済み研修医一覧の取得 |
| POST | `/api/upload_both` | Sheet1+Sheet2 一括アップロード |
| POST | `/api/run` | ソルバー実行 |
| POST | `/api/manual_assign` | 手動配置 |
| POST | `/api/manual_move` | 配置の移動 |
| POST | `/api/manual_unassign` | 配置の解除 |
| GET | `/api/download` | Excelファイルのダウンロード |
