from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from docxtpl import DocxTemplate
from datetime import datetime
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from data_dict import data_dict  # ← 別ファイル data_dict.py を参照

app = Flask(__name__)

# 🔽 あなたのLINEチャネル情報
LINE_CHANNEL_ACCESS_TOKEN ='JLHxkWqodOnZUYjdekyGfVPGecu8/QbV3v9b3/9v3QUVBt1e2VVa9iYEtjlZfyryyZ94VzBEFVjDVHhiifQybVHEgd/9G1YTyXNtpRYKYlS84prGTlQ9OEtjbYRQ0i+1Ew/LYVBKL/gOO8o28qXUNgdB04t89/1O/w1cDnyilFU='
LINE_CHANNEL_SECRET = 'c30c9f9ecce29c412c0f912f56609edd'

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SAVE_DIR = "/tmp/reports"  # Render用。ローカルなら "./reports"
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
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text="テンプレート報告書を作成中です...")
    )
    save_to_word(data_dict)  # dictからWordを生成

def save_to_word(data_dict):
    now = datetime.now()
    filename = f"report_{now.strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join(SAVE_DIR, filename)
    template_path = "来店報告書テンプレ.docx"  # テンプレート名

    try:
        print(f"📄 テンプレートを読み込み中: {template_path}")
        doc = DocxTemplate(template_path)
        doc.render(data_dict)
        doc.save(output_path)
        print(f"✅ Wordファイルを保存しました: {output_path}")

        upload_to_drive(output_path, filename)  # ← Driveへアップロード
    except Exception as e:
        print(f"❌ Wordファイル作成中にエラー発生: {e}")

def upload_to_drive(filepath, filename):
    print(f"🚀 Google Drive へアップロード開始: {filepath}")
    credentials_info = json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON"))

    credentials = service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=["https://www.googleapis.com/auth/drive.file"]
    )

    service = build("drive", "v3", credentials=credentials)
    file_metadata = {
        "name": filename,
        "parents": ["1TzWC2J5JBJXx4nr7Uu5nSHg-HUnQvh0v"]  # ← アップロード先のフォルダIDを入れる
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

    print(f"✅ Google Drive にアップロードされました (ID: {uploaded.get('id')})")
    print(f"✅ アップロード完了: https://drive.google.com/drive/u/0/my-drive")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
