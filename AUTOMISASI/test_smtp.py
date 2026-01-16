import smtplib
from dotenv import load_dotenv
import os

load_dotenv()
print('Testing SMTP connection...')
print(f'Server: {os.getenv("SMTP_SERVER")}')
print(f'Port: {os.getenv("SMTP_PORT")}')
print(f'Email: {os.getenv("SENDER_EMAIL")}')

try:
    server = smtplib.SMTP(os.getenv('SMTP_SERVER'), int(os.getenv('SMTP_PORT')))
    server.starttls()
    server.login(os.getenv('SENDER_EMAIL'), os.getenv('SENDER_PASSWORD'))
    print('✓ LOGIN BERHASIL!')
    server.quit()
except Exception as e:
    print(f'✗ GAGAL: {e}')
