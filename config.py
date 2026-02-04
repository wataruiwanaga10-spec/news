"""
ニュース収集・メール送信プログラムの設定ファイル

環境変数または直接編集で設定を行ってください。
"""
import os
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class EmailConfig:
    """メール送信の設定"""
    smtp_server: str = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port: int = int(os.getenv("SMTP_PORT", "587"))
    username: str = os.getenv("EMAIL_USERNAME", "")
    password: str = os.getenv("EMAIL_PASSWORD", "")  # Gmailの場合はアプリパスワード
    from_address: str = os.getenv("EMAIL_FROM", "")
    to_addresses: list[str] = field(default_factory=lambda: os.getenv("EMAIL_TO", "").split(","))


@dataclass
class NewsSource:
    """ニュースソースの定義"""
    name: str
    url: str
    category: str
    enabled: bool = True


# ニュースソース（RSSフィード）の設定
NEWS_SOURCES: list[NewsSource] = [
    # 日本経済ニュース
    NewsSource(
        name="Yahoo!ニュース - 経済",
        url="https://news.yahoo.co.jp/rss/topics/business.xml",
        category="経済",
    ),
    NewsSource(
        name="Yahoo!ニュース - 国際",
        url="https://news.yahoo.co.jp/rss/topics/world.xml",
        category="国際",
    ),
    NewsSource(
        name="NHK - 主要ニュース",
        url="https://www.nhk.or.jp/rss/news/cat0.xml",
        category="総合",
    ),
    NewsSource(
        name="ロイター - ビジネス",
        url="https://assets.wor.jp/rss/rdf/reuters/business.rdf",
        category="経済",
    ),
    # 経済指標・マーケット
    NewsSource(
        name="Yahoo!ファイナンス - マーケット",
        url="https://finance.yahoo.co.jp/rss/news",
        category="マーケット",
    ),
]

# 経済指標API設定（無料のAPIを使用）
MARKET_DATA_CONFIG = {
    # Yahoo Finance経由で取得する指標
    "indices": [
        {"symbol": "^N225", "name": "日経平均株価"},
        {"symbol": "^TOPX", "name": "TOPIX"},
        {"symbol": "^DJI", "name": "NYダウ"},
        {"symbol": "^GSPC", "name": "S&P 500"},
        {"symbol": "^IXIC", "name": "NASDAQ"},
    ],
    "forex": [
        {"symbol": "USDJPY=X", "name": "米ドル/円"},
        {"symbol": "EURJPY=X", "name": "ユーロ/円"},
        {"symbol": "GBPJPY=X", "name": "英ポンド/円"},
    ],
}

# 収集するニュースの件数上限（ソースごと）
MAX_NEWS_PER_SOURCE = 5

# メール送信時間（cron設定用の参考）
# 毎朝7時に送信する場合: 0 7 * * *
SEND_TIME = "07:00"
