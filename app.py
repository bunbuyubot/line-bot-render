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

app = Flask(__name__)

# 🔐 LINEチャネル情報
LINE_CHANNEL_ACCESS_TOKEN = 'JLHxkWqodOnZUYjdekyGfVPGecu8/QbV3v9b3/9v3QUVBt1e2VVa9iYEtjlZfyryyZ94VzBEFVjDVHhiifQybVHEgd/9G1YTyXNtpRYKYlS84prGTlQ9OEtjbYRQ0i+1Ew/LYVBKL/gOO8o28qXUNgdB04t89/1O/w1cDnyilFU='
LINE_CHANNEL_SECRET = 'c30c9f9ecce29c412c0f912f56609edd'

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 📁 保存先ディレクトリ（Render用）
SAVE_DIR = "/tmp/reports"
os.makedirs(SAVE_DIR, exist_ok=True)

# 📌 webhookルート
@app.route("/webhook", methods=['POST'])
def webhook():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# 🧠 メッセージを辞書に変換する
def parse_message_to_dict(message_text):
    lines = message_text.split('\n')
    data = {}
    for line in lines:
        if ':' in line:
            key, value = line.split(':', 1)
            data[key.strip()] = value.strip()
    return data

# 📩 LINEメッセージ処理
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    print("🟢 handle_message() が呼ばれました")
    text = event.message.text

    # もとのテンプレート用のdata_dictを複製
    from copy import deepcopy
    from data_dict import data_dict as base_dict
    updated_dict = deepcopy(base_dict)

    # メッセージをパース（「キー: 値」の形式で1行ずつ）
    for line in text.splitlines():
        if ':' in line:
            key, value = line.split(':', 1)
            key = key.strip()
            value = value.strip()
            if key in updated_dict:
                updated_dict[key] = value

    # 応答（非同期処理中の案内）
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text="テンプレート報告書を作成中です…")
    )

    # 置き換えたdictを元に保存
    save_to_word(updated_dict)



# 📄 Wordファイル作成
def save_to_word(data_dict):
    now = datetime.now()
    filename = f"report_{now.strftime('%Y%m%d_%H%M%S')}.docx"
    output_path = os.path.join(SAVE_DIR, filename)
    template_path = "来店報告書テンプレ.docx"

    try:
        print(f"📄 テンプレート読み込み: {template_path}")
        doc = DocxTemplate(template_path)
        doc.render(data_dict)
        doc.save(output_path)
        print(f"✅ Wordファイル保存完了: {output_path}")
        upload_to_drive(output_path, filename)
    except Exception as e:
        print(f"❌ Word作成中にエラー: {e}")

# ☁️ Google Drive にアップロード
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
        "parents": ["1TzWC2J5JBJXx4nr7Uu5nSHg-HUnQvh0v"]  # ← DriveのフォルダIDに置き換えてください
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

# ▶️ アプリ起動（Render用）
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
