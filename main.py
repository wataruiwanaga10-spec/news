#!/usr/bin/env python3
"""
ニュース収集・メール送信メインスクリプト

毎朝のニュースと経済指標を収集し、メールで送信します。
cronで定期実行することを想定しています。

使用方法:
    python main.py              # ニュース収集＋メール送信
    python main.py --dry-run    # メール送信せずに内容を確認
    python main.py --news-only  # ニュースのみ表示（市場データなし）
"""
import argparse
import logging
import sys
from datetime import datetime

from config import EmailConfig
from email_sender import EmailSender, create_html_email, create_plain_email
from news_collector import (
    MarketDataCollector,
    NewsCollector,
    format_market_data,
    format_news,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser(
        description="ニュース収集・メール送信プログラム",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="メールを送信せず、内容を標準出力に表示",
    )
    parser.add_argument(
        "--news-only",
        action="store_true",
        help="ニュースのみ収集（市場データをスキップ）",
    )
    parser.add_argument(
        "--market-only",
        action="store_true",
        help="市場データのみ収集（ニュースをスキップ）",
    )
    parser.add_argument(
        "--quiet",
        "-q",
        action="store_true",
        help="ログ出力を抑制",
    )
    args = parser.parse_args()

    if args.quiet:
        logging.getLogger().setLevel(logging.WARNING)

    now = datetime.now()
    logger.info(f"Starting news aggregation at {now.strftime('%Y-%m-%d %H:%M:%S')}")

    # 市場データ収集
    market_data_text = ""
    if not args.news_only:
        logger.info("Collecting market data...")
        try:
            market_collector = MarketDataCollector()
            market_data = market_collector.collect_all()
            market_data_text = format_market_data(market_data)
            logger.info("Market data collection completed")
        except Exception as e:
            logger.error(f"Failed to collect market data: {e}")
            market_data_text = "市場データの取得に失敗しました。"

    # ニュース収集
    news_text = ""
    if not args.market_only:
        logger.info("Collecting news...")
        try:
            news_collector = NewsCollector()
            news = news_collector.collect_all_news()
            news_text = format_news(news)
            logger.info("News collection completed")
        except Exception as e:
            logger.error(f"Failed to collect news: {e}")
            news_text = "ニュースの取得に失敗しました。"

    # メール本文作成
    plain_body = create_plain_email(market_data_text, news_text, now)
    html_body = create_html_email(market_data_text, news_text, now)
    subject = f"朝のニュースダイジェスト - {now.strftime('%Y/%m/%d')}"

    if args.dry_run:
        # ドライラン：内容を標準出力に表示
        print("=" * 80)
        print("DRY RUN MODE - メールは送信されません")
        print("=" * 80)
        print(f"\n件名: {subject}\n")
        print(plain_body)
        print("=" * 80)
        print("メール送信をスキップしました（--dry-run）")
        return 0

    # メール送信
    logger.info("Sending email...")
    sender = EmailSender(EmailConfig())

    if sender.send_email(subject, plain_body, html_body):
        logger.info("Email sent successfully!")
        return 0
    else:
        logger.error("Failed to send email")
        return 1


if __name__ == "__main__":
    sys.exit(main())
