# autorun order v1.0.6 (Update: Attendance System to logs_เช็คชื่อ)
from flask import Flask, request, abort, jsonify
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
from datetime import datetime, timedelta, timezone

app = Flask(__name__)

# ==========================================
# 🔑 1. กุญแจการเชื่อมต่อ & ตั้งค่ารูปภาพ
# ==========================================
line_bot_api = LineBotApi('msJOEakMwWxKn47C/KxhCoTAsomDcRDMK42aYVIzsuUhdaTFiLcWFBgUxbuUZtCsCL974XM/ftwTaDmS5ykI/AwmOUoVq43plbGJelanbLSb0ty5NB8rWNO+qDso2LpFU2C2Q4pDknV/eX2C9DgaMgdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('738e20aeda95ff8d4037bfe193b1626f')

cat_image_url = "https://i.postimg.cc/T2qDgNdr/87255.jpg"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet_id = "1AX07eQB9tE7dwQrnN44DS2_CgUtQ7JzjRkMgHzOP0Y4"

# เชื่อมต่อแผ่นงาน (Tabs)
sheet_curtain = client.open_by_key(sheet_id).worksheet("ผ้าม่าน")
sheet_glass = client.open_by_key(sheet_id).worksheet("งานกระจก")
sheet_attendance = client.open_by_key(sheet_id).worksheet("logs_เช็คชื่อ") # 👈 แก้ชื่อแท็บให้ตรงกับที่คุณสร้างไว้แล้ว

THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

# ==========================================
# 🧠 2. สมองกลส่วนกลาง (ผ้าม่าน/กระจก)
# ==========================================
def generate_real_order(line_order, last_real_order, is_new_year=False):
    prefix_match = re.match(r"([ก-๙a-zA-Z]+)/", line_order)
    new_prefix = prefix_match.group(1) if prefix_match else "ไม่ระบุ"
    if is_new_year or not last_real_order or "/" not in last_real_order:
        new_num = 1
    else:
        num_match = re.search(r"/(\d+)", last_real_order)
        new_num = int(num_match.group(1)) + 1 if num_match else 1
    return f"{new_prefix}/{new_num:03d}"

def process_glass_order(msg, last_real_order, last_date_in_sheet):
    date_match = re.search(r"วันที่\s*:\s*(\d+)/(\d+)/(\d+)", msg)
    if not date_match: return None
    new_d, new_m, new_y = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
    
    notification = ""
    is_new_year = False
    if last_date_in_sheet:
        old_date_parts = last_date_in_sheet.split('/')
        if len(old_date_parts) == 3:
            old_m = int(old_date_parts[1])
            old_y = int(old_date_parts[2])
            full_new_y = new_y + 2500 if new_y < 1000 else new_y
            full_old_y = old_y + 2500 if old_y < 1000 else old_y
            if full_new_y > full_old_y:
                is_new_year = True
                notification = f"📢 สมุดรันออเดอร์งานกระจกบันทึกครบของปี {full_old_y} แล้วนะคะ แนะนำให้บันทึกข้อมูลลงคอม พร้อมลบข้อมูลในชีทของปีก่อนหน้าด้วยค่ะ"
            elif new_m != old_m:
                month_name = THAI_MONTHS[old_m] if old_m <= 12 else str(old_m)
                notification = f"📢 สมุดรันออเดอร์งานกระจก เดือน:{month_name} ปี {full_old_y} เสร็จเรียบร้อยแล้วนะคะ แนะนำให้บันทึกลงคอมได้เลย"

    data = {"เลขที่ออเดอร์": "", "ชื่อลูกค้า": "", "เบอร์โทร": "", "งาน": "", "ลายละเอียด": "", "ราคา": ""}
    for line in msg.split('\n'):
        line = line.strip()
        if "เลขที่ออเดอร์" in line: data["เลขที่ออเดอร์"] = line.split(":", 1)[1].strip()
        elif "📌ชื่อ-ที่อยู่ลูกค้า" in line: data["ชื่อลูกค้า"] = line.split(":", 1)[1].strip()
        elif "📌เบอร์โทร" in line: data["เบอร์โทร"] = line.split(":", 1)[1].strip()
        elif "⭐งาน" in line: data["งาน"] = line.split(":", 1)[1].strip()
        elif "⭐ลายละเอียด" in line: data["ลายละเอียด"] = line.split(":", 1)[1].strip()
        elif "⭐ราคา" in line: data["ราคา"] = line.split(":", 1)[1].strip()

    line_order = data["เลขที่ออเดอร์"]
    search_text = data["งาน"] + " " + data["ลายละเอียด"]
    if "มุ้ง" in search_text: source = "มุ้ง"
    elif "แอร์" in search_text: source = "แอร์"
    elif "กระจก" in search_text: source = "กระจก"
    else: source = "ติดตั้ง"

    item_detail = data["ลายละเอียด"] if data["ลายละเอียด"] else data["งาน"]
    raw_total = data["ราคา"].replace(",", "").strip()
    clean_total = re.sub(r'[^\d.]', '', raw_total)
    total_display = float(clean_total) if clean_total else ""

    real_order = generate_real_order(line_order, last_real_order, is_new_year)

    final_row = [
        f"{new_d}/{new_m}/{new_y}", line_order, real_order, data["ชื่อลูกค้า"], data["เบอร์โทร"], 
        source, item_detail, total_display, "รอตรวจสอบ", "", ""
    ]
    return final_row, notification

def process_curtain_order(msg, last_real_order, last_date_in_sheet):
    date_match = re.search(r"วันที่\s*:\s*(\d+)/(\d+)/(\d+)", msg)
    if not date_match: return None
    new_d, new_m, new_y = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
    
    notification = ""
    is_new_year = False
    if last_date_in_sheet:
        old_date_parts = last_date_in_sheet.split('/')
        if len(old_date_parts) == 3:
            old_m = int(old_date_parts[1])
            old_y = int(old_date_parts[2])
            full_new_y = new_y + 2500 if new_y < 1000 else new_y
            full_old_y = old_y + 2500 if old_y < 1000 else old_y
            if full_new_y > full_old_y:
                is_new_year = True
                notification = f"📢 สมุดรันออเดอร์ผ้าม่านบันทึกครบของปี {full_old_y} แล้วนะคะ แนะนำให้บันทึกข้อมูลลงคอม พร้อมลบข้อมูลในชีทของปีก่อนหน้าด้วยค่ะ"
            elif new_m != old_m:
                month_name = THAI_MONTHS[old_m] if old_m <= 12 else str(old_m)
                notification = f"📢 สมุดรันออเดอร์ผ้าม่าน เดือน:{month_name} ปี {full_old_y} เสร็จเรียบร้อยแล้วนะคะ แนะนำให้บันทึกลงคอมได้เลย"

    if "⏬⏬⏬⏬" in msg:
        parts = msg.split("⏬⏬⏬⏬")
        top_part, bottom_part = parts[0].strip(), parts[1].strip() if len(parts) > 1 else ""
    else:
        lines = msg.split('\n')
        top_lines, bottom_lines = [], []
        found_total = False
        for line in lines:
            if found_total: bottom_lines.append(line)
            else:
                top_lines.append(line)
                if "ยอดรวม" in line: found_total = True
        top_part = '\n'.join(top_lines)
        bottom_part = '\n'.join(bottom_lines)

    data = {"วันที่": f"{new_d}/{new_m}/{new_y}", "เลขที่ออเดอร์": "", "ชื่อลูกค้า": "", "เบอร์โทร": "", "ที่มา": "", "ขนส่ง": "", "บิล": "", "ยอดรวม": ""}
    for line in top_part.split('\n'):
        if ":" in line:
            key, val = line.split(":", 1)
            clean_key = key.strip().replace("📌", "").replace("📞", "").replace("☀️", "").replace("👉", "").replace("🟥", "").replace("💳", "").replace("💰", "")
            for k in data.keys():
                if k in clean_key: data[k] = val.strip(); break

    line_order = data["เลขที่ออเดอร์"]
    source = data["ที่มา"] or ("ออนไลน์" if "อล" in line_order else "หน้าร้าน" if "นร" in line_order else "ตัวแทน" if "ตท" in line_order else "")
    transport = data["ขนส่ง"] or ("ติดตั้ง" if "นร" in line_order else "")
    bill = data["บิล"].strip().upper() or "รอชำระ"
    
    raw_total = data["ยอดรวม"].replace(",", "").strip()
    clean_total = re.sub(r'[^\d.]', '', raw_total)
    total_display = float(clean_total) if clean_total else ""

    items_lines = [l.strip() for l in bottom_part.split('\n') if l.strip()]
    short_item = items_lines[0] if items_lines else "ไม่ระบุรายการ"
    qty_match = re.search(r'(\d+)\s*(ผืน|ชุด|ชิ้น)', bottom_part)
    quantity = qty_match.group(1) if qty_match else "1"

    real_order = generate_real_order(line_order, last_real_order, is_new_year)

    final_row = [data["วันที่"], line_order, real_order, data["ชื่อลูกค้า"], data["เบอร์โทร"], source, transport, short_item, quantity, total_display, bill, ""]
    return final_row, notification

# ==========================================
# 📡 3. ท่อรับส่งข้อมูล (Webhook)
# ==========================================

@app.route("/", methods=['GET'])
def home():
    return "Bot is running 24/7!"

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try: handler.handle(body, signature)
    except InvalidSignatureError: abort(400)
    return 'OK'

# 🚪 ประตูหลังบ้านรับข้อมูลสแกนนิ้วจากคอมออฟฟิศ
@app.route("/attendance", methods=['POST'])
def receive_attendance():
    try:
        data = request.json
        name = data.get('name', 'ไม่ทราบชื่อ')
        timestamp = data.get('timestamp') # format: "2026-04-08 14:10:31"

        # แปลงเวลาให้เป็นรูปแบบไทย
        dt = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
        date_str = f"{dt.day:02d}/{dt.month:02d}/{dt.year + 543}"
        time_str = f"{dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}"

        # บันทึกลง Google Sheet แท็บ "logs_เช็คชื่อ"
        final_row = [date_str, time_str, name, "สแกนสำเร็จ"]
        sheet_attendance.append_row(final_row, value_input_option='USER_ENTERED')
        
        print(f"✅ ลงชีทสำเร็จ: {name} | เวลา: {time_str}")
        return jsonify({"status": "success"}), 200
    except Exception as e:
        print(f"❌ Error attendance: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_text = event.message.text.strip()
    tz_th = timezone(timedelta(hours=7))
    timestamp = datetime.now(tz_th).strftime('%d/%m/%Y %H:%M:%S')

    try:
        # 🔑 คำสั่งพิเศษสำหรับแอดมิน เพื่อหา Group ID เอาไว้ทำแจ้งเตือน
        if user_text == "ขอไอดีกลุ่ม":
            source_id = event.source.group_id if event.source.type == 'group' else event.source.user_id
            reply_text = f"⚙️ ไอดีของห้องแชทนี้คือ:\n{source_id}\n\n(ก๊อปปี้รหัสนี้ไปให้โปรแกรมเมอร์ได้เลยครับ)"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
            return

        # ====== 🟦 งานกระจก ======
        if "🧧🧧🧧🧧🧧" in user_text and "เลขที่ออเดอร์" in user_text:
            last_row = sheet_glass.get_all_values()[-1] if len(sheet_glass.get_all_values()) > 1 else None
            last_date, last_order = (last_row[0], last_row[2]) if last_row else (None, None)
            result = process_glass_order(user_text, last_order, last_date)
            if result:
                final_row, notification = result
                final_row.append(timestamp)
                sheet_glass.append_row(final_row, value_input_option='USER_ENTERED')
                reply_messages = [TextSendMessage(text=notification)] if notification else []
                reply_messages.append(ImageSendMessage(original_content_url=cat_image_url, preview_image_url=cat_image_url))
                line_bot_api.reply_message(event.reply_token, reply_messages)

        # ====== 🟩 งานผ้าม่าน ======
        elif "เลขที่ออเดอร์ :" in user_text:
            last_row = sheet_curtain.get_all_values()[-1] if len(sheet_curtain.get_all_values()) > 1 else None
            last_date, last_order = (last_row[0], last_row[2]) if last_row else (None, None)
            result = process_curtain_order(user_text, last_order, last_date)
            if result:
                final_row, notification = result
                final_row.append(timestamp)
                sheet_curtain.append_row(final_row, value_input_option='USER_ENTERED')
                reply_messages = [TextSendMessage(text=notification)] if notification else []
                reply_messages.append(ImageSendMessage(original_content_url=cat_image_url, preview_image_url=cat_image_url))
                line_bot_api.reply_message(event.reply_token, reply_messages)

    except Exception as e:
        print(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    app.run(port=5000)
