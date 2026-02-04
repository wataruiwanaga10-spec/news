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

    # =========================================================================
    # 九州ニュース
    # =========================================================================
    NewsSource(
        name="NHK福岡 - 九州沖縄",
        url="https://www.nhk.or.jp/rss/news/cat7.xml",
        category="九州",
    ),
    NewsSource(
        name="西日本新聞",
        url="https://www.nishinippon.co.jp/feed/",
        category="九州",
    ),
    NewsSource(
        name="RKB毎日放送",
        url="https://rkb.jp/news/rss.xml",
        category="九州",
    ),
    NewsSource(
        name="KBC九州朝日放送",
        url="https://kbc.co.jp/news/rss.xml",
        category="九州",
    ),
    NewsSource(
        name="テレQ（TVQ九州放送）",
        url="https://www.tvq.co.jp/rss/news.xml",
        category="九州",
    ),

    # --- 佐賀県 ---
    NewsSource(
        name="佐賀新聞",
        url="https://www.saga-s.co.jp/rss/news.xml",
        category="佐賀",
    ),
    NewsSource(
        name="サガテレビ（STS）",
        url="https://www.sagatv.co.jp/rss/news.xml",
        category="佐賀",
    ),
    NewsSource(
        name="NHK佐賀",
        url="https://www.nhk.or.jp/lnews/saga/index.rdf",
        category="佐賀",
    ),

    # --- 各県のニュース（必要に応じて有効化） ---
    # NewsSource(
    #     name="熊本日日新聞",
    #     url="https://kumanichi.com/rss/index.xml",
    #     category="九州",
    #     enabled=False,
    # ),
    # NewsSource(
    #     name="南日本新聞（鹿児島）",
    #     url="https://373news.com/rss/",
    #     category="九州",
    #     enabled=False,
    # ),
    # NewsSource(
    #     name="大分合同新聞",
    #     url="https://www.oita-press.co.jp/rss/",
    #     category="九州",
    #     enabled=False,
    # ),
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
# 毎朝6時30分に送信する場合: 30 6 * * *
SEND_TIME = "06:30"

# =============================================================================
# 時間フィルタリング設定
# =============================================================================
# 前日の指定時刻から当日の指定時刻までに配信されたニュースのみを取得します。
# これにより、過去に取得済みのニュースを再度取得することを防ぎます。
#
# 例: SEND_TIME="06:30" の場合
#     前日 06:30 〜 当日 06:29 に配信されたニュースを取得
# =============================================================================
TIME_FILTER_CONFIG = {
    # 時間フィルタリングを有効にするか
    "enabled": True,
    # フィルタリングの基準時刻（HH:MM形式）
    # この時刻を基準に24時間分のニュースを取得
    "base_time": "06:30",
    # 公開日時がないニュースも含めるか（Trueの場合、日時不明のニュースも取得）
    "include_no_date": True,
}

# =============================================================================
# キーワード設定
# =============================================================================
# 注目したいキーワードをリストで指定してください。
# これらのキーワードを含むニュースは「注目ニュース」として優先表示されます。
# 大文字・小文字は区別しません。
#
# 例: ["AI", "半導体", "日銀", "利上げ"]
# =============================================================================

WATCH_KEYWORDS: list[str] = [
    # --- 経済・金融 ---
    "日銀",
    "利上げ",
    "利下げ",
    "金利",
    "為替",
    "円安",
    "円高",
    "インフレ",
    "GDP",

    # --- テクノロジー ---
    "AI",
    "人工知能",
    "半導体",
    "生成AI",

    # --- 企業・業界 ---
    # "トヨタ",
    # "ソニー",

    # --- 国際 ---
    "米国",
    "中国",
    "FRB",

    # --- 自分の関心事項を追加 ---
    # "キーワード1",
    # "キーワード2",
]

# キーワードマッチングの設定
KEYWORD_CONFIG = {
    # キーワードが含まれていれば一致とみなす（部分一致）
    "partial_match": True,
    # 注目ニュースを先頭に表示するか
    "prioritize_matched": True,
}
