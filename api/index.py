import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from http.server import BaseHTTPRequestHandler

# --- CẤU HÌNH ---
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
SPREADSHEET_ID = "1x8uh0CovXS_cDbySvS6HVjWI7DOpSXYOvHjayKi4lgY"
TAB_NAME = "Data Trả góp"
TIMEZONE = ZoneInfo("Asia/Ho_Chi_Minh")
EMAIL_RECEIVER = "xaloenglishca@gmail.com"

def parse_date(date_str):
    if not date_str: return None
    for fmt in ("%d/%m/%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None

def send_summary_email(due_list, email_user, email_pass, email_type="due"):
    today_str = datetime.now(TIMEZONE).strftime('%d/%m/%Y')
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = EMAIL_RECEIVER

    if email_type == "reminder":
        msg['Subject'] = f"[WIRGROUP] NHẮC NHỞ TRẢ GÓP SAU 3 NGÀY - {today_str}"
        heading = f"Danh sách học viên sẽ đến hạn trả góp sau 3 ngày (hôm nay {today_str})"
    else:
        msg['Subject'] = f"[WIRGROUP] DANH SÁCH THU PHÍ NGÀY {today_str}"
        heading = f"Danh sách học viên đến hạn đóng tiền hôm nay ({today_str})"

    table_content = ""
    for item in due_list:
        table_content += f"<tr><td style='border:1px solid #ddd; padding:8px;'>{item['name']}</td>"
        table_content += f"<td style='border:1px solid #ddd; padding:8px;'>{item['label']}</td>"
        table_content += f"<td style='border:1px solid #ddd; padding:8px;'>{item['date']}</td>"
        table_content += f"<td style='border:1px solid #ddd; padding:8px;'>{item['amount']}</td></tr>"

    html = f"""
    <html><body>
        <h2>{heading}</h2>
        <table style='border-collapse: collapse; width: 100%;'>
            <tr style='background-color: #f2f2f2;'>
                <th style='border:1px solid #ddd; padding:8px;'>Tên học viên</th>
                <th style='border:1px solid #ddd; padding:8px;'>Đợt đóng</th>
                <th style='border:1px solid #ddd; padding:8px;'>Ngày hạn</th>
                <th style='border:1px solid #ddd; padding:8px;'>Số tiền</th>
            </tr>
            {table_content}
        </table>
    </body></html>
    """
    msg.attach(MIMEText(html, 'html'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_user, email_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Lỗi gửi mail: {e}")
        return False

def check_and_report():
    # Lấy thông tin từ Environment Variables
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    email_user = os.environ.get('EMAIL_USER')
    email_pass = os.environ.get('EMAIL_PASS')

    if not creds_json or not email_user or not email_pass:
        print("Thiếu biến môi trường!")
        return "Error: Missing Envs"

    # Kết nối Sheets dùng Dictionary thay vì file
    creds_dict = json.loads(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(TAB_NAME)
    data = sheet.get_all_values()
    
    today = datetime.now(TIMEZONE).date()
    reminder_list = []
    due_list = []

    for row in data[1:]:
        name = row[2] # Cột C
        # Check 3 lần đóng tiền (Index: L=11, M=12, N=13 | O=14, Q=16, R=17 | S=18, U=20, V=21)
        installments = [(row[11], row[12], row[13], "Lần 1"), (row[14], row[16], row[17], "Lần 2"), (row[18], row[20], row[21], "Lần 3")]

        for status, date_str, amount, label in installments:
            if status == "Chưa thanh toán" and date_str:
                due_date = parse_date(date_str)
                if not due_date:
                    continue
                item = {"name": name, "label": label, "date": date_str, "amount": amount}
                if due_date == today:
                    due_list.append(item)
                elif due_date == today + timedelta(days=3):
                    reminder_list.append(item)

    results = []
    if reminder_list:
        send_summary_email(reminder_list, email_user, email_pass, email_type="reminder")
        results.append(f"Nhac truoc 3 ngay: {len(reminder_list)} hoc vien")
    if due_list:
        send_summary_email(due_list, email_user, email_pass, email_type="due")
        results.append(f"Den han hom nay: {len(due_list)} hoc vien")

    if results:
        return "Da gui mail - " + "; ".join(results)
    return "Khong co hoc vien can nhac hoac den han hom nay."

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        result = check_and_report()
        self.send_response(200)
        self.send_header('Content-type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write(result.encode('utf-8'))