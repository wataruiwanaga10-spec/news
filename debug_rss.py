#!/usr/bin/env python3
"""
RSSフィードのデバッグスクリプト
九州・佐賀・福岡のニュースが取得できない原因を調査します。
"""
import feedparser
from datetime import datetime, timedelta

# テスト対象のRSSフィード
TEST_FEEDS = [
    ("NHK福岡", "https://www3.nhk.or.jp/lnews/fukuoka/index.rdf"),
    ("NHK佐賀", "https://www3.nhk.or.jp/lnews/saga/index.rdf"),
    ("西日本新聞 - 全国", "https://www.nishinippon.co.jp/rss/national.xml"),
    ("西日本新聞 - 福岡", "https://www.nishinippon.co.jp/rss/fukuoka.xml"),
    ("RKB毎日放送", "https://rkb.jp/news/rss.xml"),
    ("佐賀新聞 - 全国", "https://www.saga-s.co.jp/rss/all.xml"),
    ("佐賀新聞 - 佐賀", "https://www.saga-s.co.jp/rss/saga.xml"),
]

# 時間フィルタリングの範囲（参考用）
now = datetime.now()
base_hour, base_minute = 6, 30
end_time = now.replace(hour=base_hour, minute=base_minute - 1, second=59, microsecond=0)
start_time = end_time - timedelta(days=1) + timedelta(minutes=1)
if now < now.replace(hour=base_hour, minute=base_minute, second=0, microsecond=0):
    end_time -= timedelta(days=1)
    start_time -= timedelta(days=1)

print("=" * 70)
print("RSSフィード デバッグレポート")
print("=" * 70)
print(f"現在時刻: {now.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"時間フィルタ範囲: {start_time.strftime('%Y-%m-%d %H:%M')} - {end_time.strftime('%Y-%m-%d %H:%M')}")
print("=" * 70)

for name, url in TEST_FEEDS:
    print(f"\n--- {name} ---")
    print(f"URL: {url}")

    try:
        feed = feedparser.parse(url)

        # エラーチェック
        if feed.bozo:
            print(f"⚠ パースエラー: {feed.bozo_exception}")

        # ステータスチェック
        status = feed.get("status", "不明")
        print(f"HTTPステータス: {status}")

        # エントリ数
        entry_count = len(feed.entries)
        print(f"エントリ数: {entry_count}")

        if entry_count == 0:
            print("⚠ ニュースが0件です")
            continue

        # 最初のエントリをサンプルとして表示
        entry = feed.entries[0]
        print(f"\n最初のエントリ:")
        print(f"  タイトル: {entry.get('title', 'なし')[:50]}...")
        print(f"  リンク: {entry.get('link', 'なし')[:60]}...")

        # 公開日時の解析
        if hasattr(entry, "published_parsed") and entry.published_parsed:
            pub_dt = datetime(*entry.published_parsed[:6])
            print(f"  公開日時(parsed): {pub_dt.strftime('%Y-%m-%d %H:%M')}")

            # 時間フィルタリングチェック
            if start_time <= pub_dt <= end_time:
                print(f"  → ✓ 時間範囲内")
            else:
                print(f"  → ✗ 時間範囲外（フィルタで除外される）")
        elif hasattr(entry, "published"):
            print(f"  公開日時(raw): {entry.published}")
            print(f"  → ⚠ parsed形式なし（include_no_date=Trueなら含まれる）")
        else:
            print(f"  公開日時: なし")
            print(f"  → ⚠ 日時情報なし（include_no_date=Trueなら含まれる）")

        # 全エントリの時間範囲をチェック
        in_range = 0
        out_range = 0
        no_date = 0
        for e in feed.entries:
            if hasattr(e, "published_parsed") and e.published_parsed:
                pub_dt = datetime(*e.published_parsed[:6])
                if start_time <= pub_dt <= end_time:
                    in_range += 1
                else:
                    out_range += 1
            else:
                no_date += 1

        print(f"\n全エントリの時間分析:")
        print(f"  時間範囲内: {in_range}件")
        print(f"  時間範囲外: {out_range}件")
        print(f"  日時情報なし: {no_date}件")

    except Exception as e:
        print(f"✗ エラー: {e}")

print("\n" + "=" * 70)
print("デバッグ完了")
print("=" * 70)
