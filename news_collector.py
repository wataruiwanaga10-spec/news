"""
ニュース収集モジュール

RSSフィードからニュースを収集し、経済指標データを取得します。
"""
import logging
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Optional

import feedparser
import yfinance as yf

from config import (
    KEYWORD_CONFIG,
    MARKET_DATA_CONFIG,
    MAX_NEWS_PER_SOURCE,
    NEWS_SOURCES,
    TIME_FILTER_CONFIG,
    WATCH_KEYWORDS,
    NewsSource,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@dataclass
class NewsItem:
    """ニュース記事"""
    title: str
    link: str
    source: str
    category: str
    published: Optional[datetime] = None
    summary: Optional[str] = None
    matched_keywords: list[str] = None  # マッチしたキーワード

    def __post_init__(self):
        if self.matched_keywords is None:
            self.matched_keywords = []


@dataclass
class MarketData:
    """市場データ"""
    symbol: str
    name: str
    price: Optional[float] = None
    change: Optional[float] = None
    change_percent: Optional[float] = None
    error: Optional[str] = None


class NewsCollector:
    """ニュース収集クラス"""

    def __init__(
        self,
        sources: list[NewsSource] = None,
        keywords: list[str] = None,
        keyword_config: dict = None,
        time_filter_config: dict = None,
    ):
        self.sources = sources or NEWS_SOURCES
        self.keywords = keywords or WATCH_KEYWORDS
        self.keyword_config = keyword_config or KEYWORD_CONFIG
        self.time_filter_config = time_filter_config or TIME_FILTER_CONFIG

    def get_time_range(self) -> tuple[datetime, datetime]:
        """
        フィルタリング用の時間範囲を取得

        Returns:
            (start_time, end_time): 前日の基準時刻から当日の基準時刻-1分まで
        """
        base_time_str = self.time_filter_config.get("base_time", "06:30")
        hour, minute = map(int, base_time_str.split(":"))

        now = datetime.now()
        # 当日の基準時刻（終了時刻は1分前）
        end_time = now.replace(hour=hour, minute=minute - 1, second=59, microsecond=0)
        # 前日の基準時刻（開始時刻）
        start_time = end_time - timedelta(days=1) + timedelta(minutes=1)

        # 現在時刻が基準時刻より前の場合は、さらに1日前にずらす
        if now < now.replace(hour=hour, minute=minute, second=0, microsecond=0):
            end_time -= timedelta(days=1)
            start_time -= timedelta(days=1)

        return start_time, end_time

    def is_within_time_range(self, item: NewsItem) -> bool:
        """
        ニュースが指定された時間範囲内かチェック

        Args:
            item: チェックするニュースアイテム

        Returns:
            時間範囲内ならTrue
        """
        if not self.time_filter_config.get("enabled", True):
            return True

        if item.published is None:
            # 公開日時がない場合の処理
            return self.time_filter_config.get("include_no_date", True)

        start_time, end_time = self.get_time_range()
        return start_time <= item.published <= end_time

    def match_keywords(self, item: NewsItem) -> list[str]:
        """ニュースアイテムにキーワードがマッチするかチェック"""
        matched = []
        text = f"{item.title} {item.summary or ''}".lower()

        for keyword in self.keywords:
            keyword_lower = keyword.lower()
            if self.keyword_config.get("partial_match", True):
                # 部分一致
                if keyword_lower in text:
                    matched.append(keyword)
            else:
                # 完全一致（単語境界）
                import re
                if re.search(rf'\b{re.escape(keyword_lower)}\b', text):
                    matched.append(keyword)

        return matched

    def fetch_rss_feed(self, source: NewsSource) -> list[NewsItem]:
        """RSSフィードからニュースを取得"""
        news_items = []

        try:
            logger.info(f"Fetching RSS feed: {source.name}")
            feed = feedparser.parse(source.url)

            if feed.bozo and feed.bozo_exception:
                logger.warning(f"RSS parse warning for {source.name}: {feed.bozo_exception}")

            for entry in feed.entries[:MAX_NEWS_PER_SOURCE]:
                published = None
                if hasattr(entry, "published_parsed") and entry.published_parsed:
                    published = datetime(*entry.published_parsed[:6])

                summary = None
                if hasattr(entry, "summary"):
                    summary = entry.summary[:200] + "..." if len(entry.summary) > 200 else entry.summary

                news_items.append(
                    NewsItem(
                        title=entry.get("title", "タイトルなし"),
                        link=entry.get("link", ""),
                        source=source.name,
                        category=source.category,
                        published=published,
                        summary=summary,
                    )
                )

            logger.info(f"Fetched {len(news_items)} items from {source.name}")

        except Exception as e:
            logger.error(f"Error fetching RSS feed {source.name}: {e}")

        return news_items

    def collect_all_news(self) -> dict[str, list[NewsItem]]:
        """全ソースからニュースを収集"""
        news_by_category: dict[str, list[NewsItem]] = {}
        watched_news: list[NewsItem] = []

        # 時間フィルタリングが有効な場合、範囲をログに出力
        if self.time_filter_config.get("enabled", True):
            start_time, end_time = self.get_time_range()
            logger.info(f"Time filter: {start_time.strftime('%Y-%m-%d %H:%M')} - {end_time.strftime('%Y-%m-%d %H:%M')}")

        for source in self.sources:
            if not source.enabled:
                continue

            items = self.fetch_rss_feed(source)
            filtered_count = 0
            for item in items:
                # 時間フィルタリング
                if not self.is_within_time_range(item):
                    filtered_count += 1
                    continue

                # キーワードマッチング
                item.matched_keywords = self.match_keywords(item)

                if item.category not in news_by_category:
                    news_by_category[item.category] = []
                news_by_category[item.category].append(item)

                # 注目キーワードにマッチした場合は注目ニュースにも追加
                if item.matched_keywords:
                    watched_news.append(item)

            if filtered_count > 0:
                logger.info(f"Filtered out {filtered_count} old items from {source.name}")

        # 注目ニュースがあれば先頭に追加
        if watched_news and self.keyword_config.get("prioritize_matched", True):
            news_by_category = {"注目": watched_news, **news_by_category}

        return news_by_category


class MarketDataCollector:
    """市場データ収集クラス"""

    def __init__(self, config: dict = None):
        self.config = config or MARKET_DATA_CONFIG

    def fetch_quote(self, symbol: str, name: str) -> MarketData:
        """個別銘柄の価格を取得"""
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.fast_info

            price = info.get("lastPrice") or info.get("regularMarketPrice")
            prev_close = info.get("previousClose") or info.get("regularMarketPreviousClose")

            change = None
            change_percent = None
            if price and prev_close:
                change = price - prev_close
                change_percent = (change / prev_close) * 100

            return MarketData(
                symbol=symbol,
                name=name,
                price=price,
                change=change,
                change_percent=change_percent,
            )

        except Exception as e:
            logger.error(f"Error fetching quote for {symbol}: {e}")
            return MarketData(
                symbol=symbol,
                name=name,
                error=str(e),
            )

    def collect_indices(self) -> list[MarketData]:
        """株価指数を取得"""
        results = []
        for idx in self.config.get("indices", []):
            data = self.fetch_quote(idx["symbol"], idx["name"])
            results.append(data)
        return results

    def collect_forex(self) -> list[MarketData]:
        """為替レートを取得"""
        results = []
        for fx in self.config.get("forex", []):
            data = self.fetch_quote(fx["symbol"], fx["name"])
            results.append(data)
        return results

    def collect_all(self) -> dict[str, list[MarketData]]:
        """全市場データを収集"""
        return {
            "indices": self.collect_indices(),
            "forex": self.collect_forex(),
        }


def format_market_data(market_data: dict[str, list[MarketData]]) -> str:
    """市場データをテキスト形式でフォーマット"""
    lines = []

    lines.append("=" * 50)
    lines.append("株価指数")
    lines.append("=" * 50)
    for data in market_data.get("indices", []):
        if data.error:
            lines.append(f"  {data.name}: 取得エラー")
        elif data.price:
            change_str = ""
            if data.change is not None and data.change_percent is not None:
                sign = "+" if data.change >= 0 else ""
                change_str = f" ({sign}{data.change:.2f}, {sign}{data.change_percent:.2f}%)"
            lines.append(f"  {data.name}: {data.price:,.2f}{change_str}")
        else:
            lines.append(f"  {data.name}: データなし")

    lines.append("")
    lines.append("=" * 50)
    lines.append("為替レート")
    lines.append("=" * 50)
    for data in market_data.get("forex", []):
        if data.error:
            lines.append(f"  {data.name}: 取得エラー")
        elif data.price:
            change_str = ""
            if data.change is not None and data.change_percent is not None:
                sign = "+" if data.change >= 0 else ""
                change_str = f" ({sign}{data.change:.4f}, {sign}{data.change_percent:.2f}%)"
            lines.append(f"  {data.name}: {data.price:.4f}{change_str}")
        else:
            lines.append(f"  {data.name}: データなし")

    return "\n".join(lines)


def format_news(news_by_category: dict[str, list[NewsItem]]) -> str:
    """ニュースをテキスト形式でフォーマット"""
    lines = []

    # 注目ニュースを先頭に表示するため、カテゴリをソート
    categories = list(news_by_category.keys())
    if "注目" in categories:
        categories.remove("注目")
        categories = ["注目"] + sorted(categories)
    else:
        categories = sorted(categories)

    for category in categories:
        items = news_by_category[category]
        lines.append("")
        lines.append("=" * 50)
        if category == "注目":
            lines.append("★ 注目ニュース ★")
        else:
            lines.append(f"{category}ニュース")
        lines.append("=" * 50)

        for item in items:
            lines.append(f"\n■ {item.title}")
            if item.matched_keywords:
                lines.append(f"  【キーワード: {', '.join(item.matched_keywords)}】")
            lines.append(f"  出典: {item.source}")
            if item.published:
                lines.append(f"  日時: {item.published.strftime('%Y-%m-%d %H:%M')}")
            if item.summary:
                lines.append(f"  概要: {item.summary}")
            lines.append(f"  URL: {item.link}")

    return "\n".join(lines)


if __name__ == "__main__":
    # テスト実行
    print("ニュース収集テスト")
    print("-" * 50)

    news_collector = NewsCollector()
    news = news_collector.collect_all_news()
    print(format_news(news))

    print("\n市場データ収集テスト")
    print("-" * 50)

    market_collector = MarketDataCollector()
    market_data = market_collector.collect_all()
    print(format_market_data(market_data))
