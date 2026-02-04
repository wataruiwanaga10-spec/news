# News Aggregator - 朝のニュースダイジェスト

毎朝の経済指標やニュースサイトを巡回してニュースを収集し、メールで送信するプログラムです。

## 機能

- **ニュース収集**: 複数のRSSフィードから最新ニュースを収集
  - Yahoo!ニュース（経済・国際）
  - NHKニュース
  - ロイター（ビジネス）
  - Yahoo!ファイナンス

- **市場データ**: Yahoo Finance経由で経済指標を取得
  - 株価指数（日経平均、TOPIX、NYダウ、S&P500、NASDAQ）
  - 為替レート（USD/JPY、EUR/JPY、GBP/JPY）

- **メール送信**: SMTP経由でHTML/テキストメールを送信

## セットアップ

### 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 2. 環境変数の設定

メール送信に必要な環境変数を設定してください：

```bash
export SMTP_SERVER="smtp.gmail.com"       # SMTPサーバー（デフォルト: Gmail）
export SMTP_PORT="587"                     # SMTPポート（デフォルト: 587）
export EMAIL_USERNAME="your@gmail.com"     # SMTPユーザー名
export EMAIL_PASSWORD="your-app-password"  # SMTPパスワード（Gmailはアプリパスワード）
export EMAIL_FROM="your@gmail.com"         # 送信元アドレス
export EMAIL_TO="recipient@example.com"    # 送信先アドレス（カンマ区切りで複数可）
```

#### Gmailを使用する場合

1. Googleアカウントで2段階認証を有効化
2. [アプリパスワード](https://myaccount.google.com/apppasswords)を生成
3. 生成したアプリパスワードを`EMAIL_PASSWORD`に設定

## 使い方

### 基本的な使い方

```bash
# ニュース収集＋メール送信
python main.py

# ドライラン（メール送信せずに内容を確認）
python main.py --dry-run

# ニュースのみ収集（市場データなし）
python main.py --news-only

# 市場データのみ収集（ニュースなし）
python main.py --market-only

# ログ出力を抑制
python main.py --quiet
```

### cronで毎朝実行する場合

```bash
# crontabを編集
crontab -e

# 毎朝7時に実行（例）
0 7 * * * cd /path/to/news && /usr/bin/python3 main.py >> /var/log/news-digest.log 2>&1
```

環境変数を設定する場合：

```bash
0 7 * * * cd /path/to/news && EMAIL_USERNAME=xxx EMAIL_PASSWORD=xxx EMAIL_FROM=xxx EMAIL_TO=xxx /usr/bin/python3 main.py
```

または、`.env`ファイルを使用する場合は`python-dotenv`を追加してください。

## 設定のカスタマイズ

`config.py`を編集してニュースソースや市場データの設定を変更できます：

- `NEWS_SOURCES`: RSSフィードの追加・削除
- `MARKET_DATA_CONFIG`: 株価指数・為替の追加・削除
- `MAX_NEWS_PER_SOURCE`: ソースごとのニュース件数上限

## ファイル構成

```
news/
├── main.py            # メインスクリプト
├── config.py          # 設定ファイル
├── news_collector.py  # ニュース・市場データ収集
├── email_sender.py    # メール送信
├── requirements.txt   # 依存パッケージ
└── README.md          # このファイル
```

## トラブルシューティング

### メールが送信できない

1. 環境変数が正しく設定されているか確認
2. Gmailの場合、アプリパスワードを使用しているか確認
3. ファイアウォールでSMTPポートがブロックされていないか確認

### ニュースが取得できない

1. インターネット接続を確認
2. RSSフィードのURLが有効か確認
3. `python news_collector.py`で個別にテスト

### 市場データが取得できない

1. `yfinance`がインストールされているか確認
2. Yahoo Financeのレート制限に引っかかっていないか確認
