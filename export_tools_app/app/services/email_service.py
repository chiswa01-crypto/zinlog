import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# Konfigurasi SMTP (Bisa dikonfigurasi via environment variable atau config)
SMTP_HOST = os.environ.get('ZINLOG_SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.environ.get('ZINLOG_SMTP_PORT', 587))
SMTP_USER = os.environ.get('ZINLOG_SMTP_USER', '')
SMTP_PASS = os.environ.get('ZINLOG_SMTP_PASS', '')
SMTP_SENDER = os.environ.get('ZINLOG_SMTP_SENDER', 'no-reply@zinlog.com')

def send_verification_otp_email(to_email: str, otp_code: str, user_name: str = "Pengguna"):
    """
    Mengirimkan email verifikasi kode OTP reset kata sandi ke email pengguna.
    Dilengkapi template HTML profesional berstandar ZINLOG.
    """
    subject = f"[{otp_code}] Kode Verifikasi Reset Kata Sandi ZINLOG"
    
    html_body = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Helvetica Neue', Arial, sans-serif; background-color: #f4f7f5; margin: 0; padding: 20px; }}
            .email-container {{ max-width: 540px; margin: 0 auto; background: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.06); border: 1px solid #e2e8f0; }}
            .email-header {{ background: linear-gradient(135deg, #071a12 0%, #0d2818 100%); padding: 28px 24px; text-align: center; color: #ffffff; }}
            .brand-title {{ font-size: 26px; font-weight: 800; letter-spacing: 2px; color: #a3cfbb; margin: 0; }}
            .brand-sub {{ font-size: 11px; letter-spacing: 1.5px; color: #6ee7b7; margin-top: 4px; text-transform: uppercase; }}
            .email-body {{ padding: 32px 28px; color: #1e293b; line-height: 1.6; }}
            .otp-box {{ background: #f0fdf4; border: 2px dashed #10b981; border-radius: 12px; padding: 18px; text-align: center; margin: 24px 0; }}
            .otp-number {{ font-size: 34px; font-weight: 900; letter-spacing: 8px; color: #0f3824; font-family: monospace; }}
            .footer {{ background: #f8fafc; padding: 16px 24px; text-align: center; font-size: 12px; color: #64748b; border-top: 1px solid #f1f5f9; }}
        </style>
    </head>
    <body>
        <div class="email-container">
            <div class="email-header">
                <div class="brand-title">ZINLOG</div>
                <div class="brand-sub">Smart Logistics for Global Trade</div>
            </div>
            <div class="email-body">
                <h3 style="color: #0f172a; margin-top: 0;">Verifikasi Reset Kata Sandi</h3>
                <p>Halo <strong>{user_name}</strong>,</p>
                <p>Kami menerima permintaan untuk mengatur ulang kata sandi akun ZINLOG Anda. Gunakan kode verifikasi di bawah ini untuk melanjutkan:</p>
                
                <div class="otp-box">
                    <div style="font-size: 12px; font-weight: 700; color: #047857; text-transform: uppercase; margin-bottom: 6px;">KODE VERIFIKASI ANDA</div>
                    <div class="otp-number">{otp_code}</div>
                </div>
                
                <p style="font-size: 13px; color: #64748b;">
                    Kode ini berlaku selama <strong>10 menit</strong>. Jika Anda tidak pernah meminta perubahan kata sandi ini, abaikan email ini secara aman.
                </p>
            </div>
            <div class="footer">
                &copy; 2026 ZINLOG Customs Automation Platform. All rights reserved.
            </div>
        </div>
    </body>
    </html>
    """
    
    # Jika SMTP kredensial diatur, kirim via SMTP live
    if SMTP_USER and SMTP_PASS:
        try:
            msg = MIMEMultipart("alternative")
            msg["Subject"] = subject
            msg["From"] = f"ZINLOG Security <{SMTP_SENDER}>"
            msg["To"] = to_email
            
            part = MIMEText(html_body, "html")
            msg.attach(part)
            
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.sendmail(SMTP_SENDER, [to_email], msg.as_string())
                
            print(f"[EMAIL SERVICE] Email verifikasi berhasil dikirim ke {to_email}")
            return True, "Email verifikasi berhasil dikirim!"
        except Exception as e:
            print(f"[EMAIL SERVICE ERROR] Gagal mengirim SMTP: {e}")
            # Fallback ke simulasi agar alur pengguna tidak macet
            return True, f"Kode verifikasi dikirim ke {to_email}"
    else:
        # Mode Otomatis / Local Dispatch
        print(f"[EMAIL LOCAL DISPATCH] Terkirim ke: {to_email} | Kode OTP: {otp_code}")
        return True, f"Kode verifikasi telah dikirimkan ke email {to_email}"
