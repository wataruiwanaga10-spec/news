"""
ニュース収集モジュール

RSSフィードからニュースを収集し、経済指標データを取得します。
"""
import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import feedparser
import yfinance as yf

from config import (
    MARKET_DATA_CONFIG,
    MAX_NEWS_PER_SOURCE,
    NEWS_SOURCES,
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

    def __init__(self, sources: list[NewsSource] = None):
        self.sources = sources or NEWS_SOURCES

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

        for source in self.sources:
            if not source.enabled:
                continue

            items = self.fetch_rss_feed(source)
            for item in items:
                if item.category not in news_by_category:
                    news_by_category[item.category] = []
                news_by_category[item.category].append(item)

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

    for category, items in sorted(news_by_category.items()):
        lines.append("")
        lines.append("=" * 50)
        lines.append(f"{category}ニュース")
        lines.append("=" * 50)

        for item in items:
            lines.append(f"\n■ {item.title}")
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
