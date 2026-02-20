import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import datetime

def send_email_report(subject, html_content, recipients):
    """
    Sends an HTML email report using SMTP credentials from environment variables.
    
    Args:
        subject (str): The subject line of the email.
        html_content (str): The HTML body of the email.
        recipients (list): List of email addresses to send to.
    """
    # 1. Load Credentials
    # 1. Load Credentials
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com").strip()
    smtp_port = int(os.getenv("SMTP_PORT", "587").strip())
    
    sender_email = os.getenv("SMTP_EMAIL", "")
    sender_password = os.getenv("SMTP_PASSWORD", "")
    
    if sender_email: sender_email = sender_email.strip()
    if sender_password: sender_password = sender_password.strip()
    
    if not sender_email or not sender_password:
        print("⚠️ Skipping Email: SMTP_EMAIL or SMTP_PASSWORD not set in .env")
        return

    # 2. Create Message
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Luxvance Bot <{sender_email}>"
    msg["To"] = ", ".join(recipients)

    # 3. Attach HTML Body
    # Add a simple style header for basic formatting
    full_html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #333; }}
            .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
            h2 {{ color: #0b333d; border-bottom: 2px solid #0b333d; padding-bottom: 10px; }}
            .stat-box {{ background: #f4f6f8; padding: 15px; border-radius: 8px; margin-bottom: 20px; }}
            .stat-row {{ display: flex; justify-content: space-between; margin-bottom: 8px; }}
            .stat-label {{ color: #666; }}
            .stat-value {{ font-weight: bold; color: #111; }}
            .footer {{ font-size: 12px; color: #999; margin-top: 30px; text-align: center; }}
        </style>
    </head>
    <body>
        <div class="container">
            {html_content}
            
            <div class="footer">
                🚀 Sent by AntiGravity Automation • {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}
            </div>
        </div>
    </body>
    </html>
    """
    
    msg.attach(MIMEText(full_html, "html"))

    # 4. Send
    try:
        print(f"📧 Connecting to SMTP ({smtp_server}:{smtp_port})...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        print(f"✅ Email sent to: {recipients}")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
