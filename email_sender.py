"""
メール送信モジュール

SMTPを使用してHTMLまたはテキストメールを送信します。
"""
import logging
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional

from config import EmailConfig

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class EmailSender:
    """メール送信クラス"""

    def __init__(self, config: EmailConfig = None):
        self.config = config or EmailConfig()

    def _validate_config(self) -> bool:
        """設定の検証"""
        if not self.config.username:
            logger.error("EMAIL_USERNAME is not set")
            return False
        if not self.config.password:
            logger.error("EMAIL_PASSWORD is not set")
            return False
        if not self.config.from_address:
            logger.error("EMAIL_FROM is not set")
            return False
        if not self.config.to_addresses or not self.config.to_addresses[0]:
            logger.error("EMAIL_TO is not set")
            return False
        return True

    def send_email(
        self,
        subject: str,
        body_text: str,
        body_html: Optional[str] = None,
    ) -> bool:
        """メールを送信"""
        if not self._validate_config():
            logger.error("Email configuration is invalid")
            return False

        try:
            # メッセージ作成
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = self.config.from_address
            msg["To"] = ", ".join(self.config.to_addresses)

            # テキストパート
            part_text = MIMEText(body_text, "plain", "utf-8")
            msg.attach(part_text)

            # HTMLパート（指定された場合）
            if body_html:
                part_html = MIMEText(body_html, "html", "utf-8")
                msg.attach(part_html)

            # SMTP接続・送信
            logger.info(f"Connecting to SMTP server: {self.config.smtp_server}:{self.config.smtp_port}")

            with smtplib.SMTP(self.config.smtp_server, self.config.smtp_port) as server:
                server.starttls()
                server.login(self.config.username, self.config.password)
                server.send_message(msg)

            logger.info(f"Email sent successfully to {self.config.to_addresses}")
            return True

        except smtplib.SMTPAuthenticationError as e:
            logger.error(f"SMTP authentication failed: {e}")
            return False
        except smtplib.SMTPException as e:
            logger.error(f"SMTP error: {e}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error sending email: {e}")
            return False


def create_html_email(
    market_data_text: str,
    news_text: str,
    date: datetime = None,
) -> str:
    """HTMLメール本文を作成"""
    if date is None:
        date = datetime.now()

    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{
            font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }}
        h1 {{
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #2980b9;
            margin-top: 30px;
            border-left: 4px solid #3498db;
            padding-left: 10px;
        }}
        .market-section {{
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
        }}
        .market-section pre {{
            margin: 0;
            white-space: pre-wrap;
            font-family: 'Monaco', 'Consolas', monospace;
            font-size: 14px;
        }}
        .news-item {{
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 15px;
            margin: 15px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }}
        .news-title {{
            font-weight: bold;
            color: #2c3e50;
            font-size: 16px;
            margin-bottom: 8px;
        }}
        .news-meta {{
            color: #7f8c8d;
            font-size: 12px;
            margin-bottom: 8px;
        }}
        .news-summary {{
            color: #555;
            font-size: 14px;
        }}
        .news-link {{
            display: inline-block;
            margin-top: 8px;
            color: #3498db;
            text-decoration: none;
        }}
        .news-link:hover {{
            text-decoration: underline;
        }}
        .footer {{
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
            color: #7f8c8d;
            font-size: 12px;
            text-align: center;
        }}
        .positive {{
            color: #27ae60;
        }}
        .negative {{
            color: #e74c3c;
        }}
    </style>
</head>
<body>
    <h1>朝のニュースダイジェスト</h1>
    <p>{date.strftime('%Y年%m月%d日 %H:%M')} 配信</p>

    <h2>市場データ</h2>
    <div class="market-section">
        <pre>{market_data_text}</pre>
    </div>

    <h2>ニュース</h2>
    <div class="news-section">
        <pre>{news_text}</pre>
    </div>

    <div class="footer">
        <p>このメールは自動送信されています。</p>
        <p>News Aggregator - Daily Digest</p>
    </div>
</body>
</html>
"""
    return html


def create_plain_email(
    market_data_text: str,
    news_text: str,
    date: datetime = None,
) -> str:
    """テキストメール本文を作成"""
    if date is None:
        date = datetime.now()

    text = f"""
================================================================================
朝のニュースダイジェスト
{date.strftime('%Y年%m月%d日 %H:%M')} 配信
================================================================================

【市場データ】
{market_data_text}

【ニュース】
{news_text}

--------------------------------------------------------------------------------
このメールは自動送信されています。
News Aggregator - Daily Digest
--------------------------------------------------------------------------------
"""
    return text


if __name__ == "__main__":
    # テスト
    print("メール送信テスト")
    print("-" * 50)

    sender = EmailSender()
    if sender._validate_config():
        print("設定OK: メール送信可能")
    else:
        print("設定エラー: 環境変数を確認してください")
        print("必要な環境変数:")
        print("  - EMAIL_USERNAME: SMTPユーザー名")
        print("  - EMAIL_PASSWORD: SMTPパスワード")
        print("  - EMAIL_FROM: 送信元アドレス")
        print("  - EMAIL_TO: 送信先アドレス（カンマ区切り）")
