
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from docxtpl import DocxTemplate, RichText
from datetime import datetime
import os
import json
from copy import deepcopy
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from data_dict import data_dict


app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")

# ❗チェック：もし読み込めてなかったら例外を出す
if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_CHANNEL_SECRET:
    raise ValueError("LINEのトークンまたはシークレットが設定されていません")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SAVE_DIR = "/tmp/reports"
os.makedirs(SAVE_DIR, exist_ok=True)

@app.route("/webhook", methods=['POST'])
def webhook():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    print("🟢 handle_message() が呼ばれました")
    print("🧪 TOKEN:", LINE_CHANNEL_ACCESS_TOKEN[:10] if LINE_CHANNEL_ACCESS_TOKEN else "None")
    print("🧪 SECRET:", LINE_CHANNEL_SECRET[:10] if LINE_CHANNEL_SECRET else "None")
    text = event.message.text

    # 🔽 ここに追記！
    print("📥 LINEメッセージ内容:", text)
    print("📤 返信トークン:", event.reply_token)

    updated_dict = deepcopy(data_dict)

    for line in text.splitlines():
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            if key in updated_dict:
                updated_dict[key] = value

    print(f"📦 更新された dict: {updated_dict}")

    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text="テンプレート報告書を作成中です...")
    )

    save_to_word(updated_dict)

def upload_to_drive(filepath, filename):
    print(f"🚀 Google Driveへアップロード開始: {filepath}")
    credentials_info = json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON"))

    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=["https://www.googleapis.com/auth/drive.file"]
    )

    service = build("drive", "v3", credentials=credentials)
    file_metadata = {
        "name": filename,
        "parents": ["1TzWC2J5JBJXx4nr7Uu5nSHg-HUnQvh0v"]
    }

    media = MediaFileUpload(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    print(f"✅ Driveアップロード完了 (ID: {uploaded.get('id')})")

def convert_newlines(value):
    rt = RichText()
    for i, line in enumerate(value.split('\n')):
        if i > 0:
            rt.add_break()
        rt.add(line)
    return rt

def save_to_word(data_dict):
    now = datetime.now()
    filename = f"report_{now.strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join(SAVE_DIR, filename)
    
    # 🔽 ここに追記！
    print("🧪 Wordファイル保存先:", output_path)

    template_path = "template.docx"
    ...

    try:
        print(f"📄 テンプレート読み込み: {template_path}")
        doc = DocxTemplate(template_path)

        fields_with_newlines = ['良い兆候', '課題', '提案', '店舗様のお言葉','稼働率']
        for field in fields_with_newlines:
            if field in data_dict and isinstance(data_dict[field], str):
                data_dict[field] = convert_newlines(data_dict[field])

        doc.render(data_dict)
        doc.save(output_path)
        print(f"✅ Wordファイル保存完了: {output_path}")

        upload_to_drive(output_path, filename)

    except Exception as e:
        print(f"❌ Word作成中にエラー: {e}")

        print("🧪 TOKEN:", LINE_CHANNEL_ACCESS_TOKEN[:10] if LINE_CHANNEL_ACCESS_TOKEN else "None")
        print("🧪 SECRET:", LINE_CHANNEL_SECRET[:10] if LINE_CHANNEL_SECRET else "None")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
