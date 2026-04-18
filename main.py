import os
import json
import tempfile
import requests
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi,
    ReplyMessageRequest, TextMessage, AudioMessage
)
from linebot.v3.webhooks import MessageEvent, AudioMessageContent, TextMessageContent
import google.generativeai as genai
from openai import OpenAI
import re
from datetime import datetime

app = Flask(__name__)

# 環境變數
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

# 初始化
handler = WebhookHandler(LINE_CHANNEL_SECRET)
configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
genai.configure(api_key=GEMINI_API_KEY)
openai_client = OpenAI(api_key=OPENAI_API_KEY)

# 暫存對話狀態（簡單的 in-memory，每次重啟會清除）
user_sessions = {}

# ─────────────────────────────────────────
# Webhook 進入點
# ─────────────────────────────────────────
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers["X-Line-Signature"]
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

# 保持伺服器喚醒
@app.route("/", methods=["GET"])
def health():
    return "Bot is running!"

# ─────────────────────────────────────────
# 處理語音訊息
# ─────────────────────────────────────────
@handler.add(MessageEvent, message=AudioMessageContent)
def handle_audio(event):
    user_id = event.source.user_id
    reply_token = event.reply_token

    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)

        # 下載語音檔
        audio_content = line_bot_api.get_message_content(event.message.id)
        with tempfile.NamedTemporaryFile(suffix=".m4a", delete=False) as f:
            for chunk in audio_content.iter_content():
                f.write(chunk)
            audio_path = f.name

        # 通知用戶處理中
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text="⏳ 語音處理中，請稍候...")]
            )
        )

        try:
            # 呼叫 Gemini 做逐字稿 + 語者辨識
            transcript, speaker_samples, meeting_info = transcribe_with_gemini(audio_path)
            os.unlink(audio_path)

            # 儲存 session
            user_sessions[user_id] = {
                "state": "waiting_speaker_confirm" if meeting_info["confirmed"] else "waiting_meeting_info",
                "transcript": transcript,
                "speaker_samples": speaker_samples,
                "meeting_info": meeting_info,
                "speaker_map": {}
            }

            if not meeting_info["confirmed"]:
                # 需要確認會議資訊
                msg = f"📋 請提供會議資訊（語音中未偵測到）：\n\n請回覆格式：\n日期=2024/01/15，名稱=Q3預算會議"
                send_push_message(user_id, msg)
            else:
                # 直接進入語者確認
                msg = build_speaker_confirm_message(speaker_samples, meeting_info)
                send_push_message(user_id, msg)

        except Exception as e:
            send_push_message(user_id, f"❌ 處理語音時發生錯誤：{str(e)}\n請重新傳送語音。")

# ─────────────────────────────────────────
# 處理文字訊息（回覆語者確認）
# ─────────────────────────────────────────
@handler.add(MessageEvent, message=TextMessageContent)
def handle_text(event):
    user_id = event.source.user_id
    reply_token = event.reply_token
    text = event.message.text.strip()

    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)

        session = user_sessions.get(user_id)

        if not session:
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=reply_token,
                    messages=[TextMessage(text="請先傳送語音檔，我會幫您整理會議記錄 🎙️")]
                )
            )
            return

        state = session.get("state")

        # ── 等待會議資訊 ──
        if state == "waiting_meeting_info":
            meeting_info = parse_meeting_info(text)
            if not meeting_info:
                line_bot_api.reply_message(
                    ReplyMessageRequest(
                        reply_token=reply_token,
                        messages=[TextMessage(text="格式不正確，請使用：\n日期=2024/01/15，名稱=Q3預算會議")]
                    )
                )
                return
            session["meeting_info"].update(meeting_info)
            session["meeting_info"]["confirmed"] = True
            session["state"] = "waiting_speaker_confirm"
            user_sessions[user_id] = session

            msg = build_speaker_confirm_message(session["speaker_samples"], session["meeting_info"])
            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=reply_token,
                    messages=[TextMessage(text=msg)]
                )
            )

        # ── 等待語者確認 ──
        elif state == "waiting_speaker_confirm":
            speaker_map = parse_speaker_map(text)
            if not speaker_map:
                line_bot_api.reply_message(
                    ReplyMessageRequest(
                        reply_token=reply_token,
                        messages=[TextMessage(text="格式不正確，請使用：\n語者1=王經理，語者2=我，語者3=陳會計")]
                    )
                )
                return

            session["speaker_map"] = speaker_map
            session["state"] = "generating"
            user_sessions[user_id] = session

            line_bot_api.reply_message(
                ReplyMessageRequest(
                    reply_token=reply_token,
                    messages=[TextMessage(text="✅ 收到！正在生成會議記錄 Word 檔，請稍候...")]
                )
            )

            # 生成 Word 文件
            try:
                word_path = generate_meeting_word(
                    session["transcript"],
                    session["speaker_map"],
                    session["meeting_info"]
                )
                # 上傳並傳送 Word 檔
                send_file_to_line(user_id, word_path, session["meeting_info"])
                os.unlink(word_path)
                del user_sessions[user_id]
            except Exception as e:
                send_push_message(user_id, f"❌ 生成 Word 時發生錯誤：{str(e)}")

# ─────────────────────────────────────────
# Gemini 逐字稿 + 語者辨識
# ─────────────────────────────────────────
def transcribe_with_gemini(audio_path):
    model = genai.GenerativeModel("gemini-1.5-pro")

    with open(audio_path, "rb") as f:
        audio_data = f.read()

    prompt = """請分析這段音訊，完成以下任務：

1. 產生完整逐字稿，格式為：
[語者N][時間戳] 發言內容

2. 嘗試從對話內容判斷：
- 會議日期（如有提及）
- 會議名稱或主題（如有提及）

3. 回傳 JSON 格式（只回傳 JSON，不要其他文字）：
{
  "transcript": "[語者1][00:00] 發言內容\n[語者2][00:30] 發言內容...",
  "speakers": ["語者1", "語者2", "語者3"],
  "meeting_date": "2024/01/15 或 null",
  "meeting_name": "會議名稱 或 null",
  "speaker_first_appearance": {
    "語者1": "00:00",
    "語者2": "00:30"
  },
  "speaker_samples": {
    "語者1": ["發言1", "發言2"],
    "語者2": ["發言1", "發言2"]
  },
  "total_duration_seconds": 600
}

語者取樣規則：
- 每位語者取前2~3句發言
- 若語者第一次出現時間超過總時長1/3，改取5句"""

    response = model.generate_content([
        {"mime_type": "audio/mp4", "data": audio_data},
        prompt
    ])

    raw = response.text.strip()
    # 清除可能的 markdown code block
    raw = re.sub(r"```json|```", "", raw).strip()
    data = json.loads(raw)

    transcript = data["transcript"]
    speaker_samples = data["speaker_samples"]
    total_seconds = data.get("total_duration_seconds", 3600)
    first_appearance = data.get("speaker_first_appearance", {})

    # 判斷晚出現的語者（超過1/3總時長）
    late_threshold = total_seconds / 3
    for speaker, samples in speaker_samples.items():
        ts = first_appearance.get(speaker, "00:00")
        seconds = timestamp_to_seconds(ts)
        if seconds > late_threshold:
            speaker_samples[speaker] = {"samples": samples, "late": True, "first_ts": ts}
        else:
            speaker_samples[speaker] = {"samples": samples, "late": False, "first_ts": ts}

    meeting_info = {
        "date": data.get("meeting_date"),
        "name": data.get("meeting_name"),
        "confirmed": bool(data.get("meeting_date") and data.get("meeting_name"))
    }

    return transcript, speaker_samples, meeting_info

def timestamp_to_seconds(ts):
    try:
        parts = ts.split(":")
        if len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        elif len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    except:
        return 0
    return 0

# ─────────────────────────────────────────
# 建立語者確認訊息
# ─────────────────────────────────────────
def build_speaker_confirm_message(speaker_samples, meeting_info):
    lines = ["🎙️ 請確認以下語者身份：\n"]

    for speaker, info in speaker_samples.items():
        samples = info["samples"]
        late = info["late"]
        ts = info["first_ts"]

        if late:
            lines.append(f"⚠️【{speaker}】（{ts} 才出現，提供較多內容供辨識）")
        else:
            lines.append(f"【{speaker}】（{ts} 出現）")

        for s in samples:
            lines.append(f"  「{s}」")
        lines.append("")

    lines.append("請回覆語者對應，例如：")
    lines.append("語者1=王經理，語者2=我，語者3=陳會計")

    if meeting_info.get("date") and meeting_info.get("name"):
        lines.append(f"\n📋 偵測到會議資訊：")
        lines.append(f"日期：{meeting_info['date']}")
        lines.append(f"名稱：{meeting_info['name']}")

    return "\n".join(lines)

# ─────────────────────────────────────────
# 解析用戶回覆
# ─────────────────────────────────────────
def parse_speaker_map(text):
    """解析 '語者1=王經理，語者2=我' 格式"""
    pattern = r"語者(\d+)\s*=\s*([^\s，,、]+)"
    matches = re.findall(pattern, text)
    if not matches:
        return None
    return {f"語者{num}": name for num, name in matches}

def parse_meeting_info(text):
    """解析 '日期=2024/01/15，名稱=Q3預算會議' 格式"""
    date_match = re.search(r"日期\s*=\s*([^\s，,、]+)", text)
    name_match = re.search(r"名稱\s*=\s*(.+?)(?:[，,]|$)", text)
    if not date_match or not name_match:
        return None
    return {
        "date": date_match.group(1).strip(),
        "name": name_match.group(1).strip()
    }

# ─────────────────────────────────────────
# GPT 生成會議記錄內容
# ─────────────────────────────────────────
def generate_meeting_content(transcript, speaker_map, meeting_info):
    # 替換語者名稱
    replaced_transcript = transcript
    for speaker_key, real_name in speaker_map.items():
        replaced_transcript = replaced_transcript.replace(speaker_key, real_name)

    meeting_prompt = f"""# 📝【會議紀錄整理模組｜逐字稿轉正式紀錄】
你是一位【專業會議記錄助理】，擅長【聽寫轉錄、摘要重點、行動項目擷取與格式化編輯】。

🔸任務說明
請整理以下逐字稿為正式會議記錄，以 JSON 格式回傳所有內容（稍後會轉成 Word）。

📋 會議基本資訊：
- 日期：{meeting_info['date']}
- 名稱：{meeting_info['name']}
- 出席人員：{', '.join(speaker_map.values())}

📝 逐字稿：
{replaced_transcript}

🎯 請回傳 JSON（只回傳 JSON，不要其他文字）：
{{
  "meeting_date": "yyyy/mm/dd",
  "meeting_name": "會議名稱",
  "attendees": ["人員1", "人員2"],
  "location": "（逐字稿未提及則填空字串）",
  "recorder": "（逐字稿未提及則填空字串）",
  "topics": [
    {{
      "title": "討論主題標題",
      "points": ["各方意見或重點1", "各方意見或重點2"]
    }}
  ],
  "action_items": [
    {{
      "category": "分類",
      "content": "行動項目內容",
      "owner": "負責人",
      "due_date": "預計完成時間",
      "notes": "備註"
    }}
  ],
  "pending_items": ["未決事項1", "未決事項2"],
  "remarks": ["補充備註1"]
}}"""

    response = openai_client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": meeting_prompt}],
        temperature=0.3
    )
    raw = response.choices[0].message.content.strip()
    raw = re.sub(r"```json|```", "", raw).strip()
    return json.loads(raw)

# ─────────────────────────────────────────
# 生成 Word 文件（Node.js docx）
# ─────────────────────────────────────────
def generate_meeting_word(transcript, speaker_map, meeting_info):
    content = generate_meeting_content(transcript, speaker_map, meeting_info)

    date = content["meeting_date"]
    name = content["meeting_name"]
    filename = f"{date.replace('/', '')}_{name}.docx"
    output_path = f"/tmp/{filename}"

    # 建立 JS 腳本
    js_script = build_docx_js(content, output_path)
    js_path = "/tmp/generate_doc.mjs"
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(js_script)

    import subprocess
    result = subprocess.run(["node", js_path], capture_output=True, text=True)
    if result.returncode != 0:
        raise Exception(f"Node.js error: {result.stderr}")

    return output_path

def build_docx_js(content, output_path):
    topics_js = ""
    for topic in content.get("topics", []):
        points_js = "\n".join([
            f'new Paragraph({{ numbering: {{ reference: "bullets", level: 0 }}, children: [new TextRun({{ text: {json.dumps(p)}, font: "Microsoft JhengHei", size: 24 }})] }}),'
            for p in topic["points"]
        ])
        topics_js += f"""
new Paragraph({{ children: [new TextRun({{ text: {json.dumps(topic["title"])}, bold: true, size: 28, font: "Microsoft JhengHei" }})], spacing: {{ before: 120, after: 60 }} }}),
{points_js}
"""

    action_rows_js = ""
    for item in content.get("action_items", []):
        def cell(text, width=1620):
            return f"""new TableCell({{
  width: {{ size: {width}, type: WidthType.DXA }},
  borders: {{ top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder }},
  margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
  children: [new Paragraph({{ children: [new TextRun({{ text: {json.dumps(text)}, font: "Microsoft JhengHei", size: 22 }})] }})]
}})"""
        action_rows_js += f"""new TableRow({{ children: [
  {cell(item.get("category",""), 1200)},
  {cell(item.get("content",""), 2800)},
  {cell(item.get("owner",""), 1400)},
  {cell(item.get("due_date",""), 1800)},
  {cell(item.get("notes",""), 1800)},
] }}),"""

    pending_js = "\n".join([
        f'new Paragraph({{ numbering: {{ reference: "bullets", level: 0 }}, children: [new TextRun({{ text: {json.dumps(p)}, font: "Microsoft JhengHei", size: 24 }})] }}),'
        for p in content.get("pending_items", [])
    ])

    remarks_js = "\n".join([
        f'new Paragraph({{ numbering: {{ reference: "bullets", level: 0 }}, children: [new TextRun({{ text: {json.dumps(r)}, font: "Microsoft JhengHei", size: 24 }})] }}),'
        for r in content.get("remarks", [])
    ])

    attendees_str = "、".join(content.get("attendees", []))

    return f"""
import {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
         AlignmentType, LevelFormat, BorderStyle, WidthType }} from 'docx';
import fs from 'fs';

const blackBorder = {{ style: BorderStyle.SINGLE, size: 4, color: "000000" }};

const doc = new Document({{
  numbering: {{
    config: [{{
      reference: "bullets",
      levels: [{{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: {{ paragraph: {{ indent: {{ left: 720, hanging: 360 }} }} }} }}]
    }}]
  }},
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 11906, height: 16838 }},
        margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }}
      }}
    }},
    children: [
      // 主標題
      new Paragraph({{
        alignment: AlignmentType.CENTER,
        spacing: {{ after: 240 }},
        children: [new TextRun({{
          text: {json.dumps(content["meeting_date"] + " " + content["meeting_name"] + "紀錄")},
          bold: true, size: 48, font: "Microsoft JhengHei"
        }})]
      }}),

      // 一、會議基本資訊
      new Paragraph({{ spacing: {{ before: 240, after: 120 }}, children: [new TextRun({{ text: "一、會議基本資訊", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "1. 會議名稱：" + {json.dumps(content["meeting_name"])}, font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "2. 會議日期：" + {json.dumps(content["meeting_date"])}, font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "3. 出席人員：" + {json.dumps(attendees_str)}, font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "4. 地點：" + {json.dumps(content.get("location",""))}, font: "Microsoft JhengHei", size: 24 }})] }}),
      new Paragraph({{ children: [new TextRun({{ text: "5. 記錄人：" + {json.dumps(content.get("recorder",""))}, font: "Microsoft JhengHei", size: 24 }})] }}),

      // 二、討論主題摘要
      new Paragraph({{ spacing: {{ before: 240, after: 120 }}, children: [new TextRun({{ text: "二、討論主題摘要與各方意見", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      {topics_js}

      // 三、決議事項與行動項目
      new Paragraph({{ spacing: {{ before: 240, after: 120 }}, children: [new TextRun({{ text: "三、決議事項與行動項目", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      new Table({{
        width: {{ size: 9000, type: WidthType.DXA }},
        columnWidths: [1200, 2800, 1400, 1800, 1800],
        rows: [
          new TableRow({{
            tableHeader: true,
            children: [
              new TableCell({{ width: {{ size: 1200, type: WidthType.DXA }}, borders: {{ top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "分類", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
              new TableCell({{ width: {{ size: 2800, type: WidthType.DXA }}, borders: {{ top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "行動項目內容", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
              new TableCell({{ width: {{ size: 1400, type: WidthType.DXA }}, borders: {{ top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "負責人", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
              new TableCell({{ width: {{ size: 1800, type: WidthType.DXA }}, borders: {{ top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "預計完成時間", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
              new TableCell({{ width: {{ size: 1800, type: WidthType.DXA }}, borders: {{ top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder }}, margins: {{ top: 80, bottom: 80, left: 120, right: 120 }}, shading: {{ fill: "D9D9D9" }}, children: [new Paragraph({{ children: [new TextRun({{ text: "備註", bold: true, font: "Microsoft JhengHei", size: 22 }})] }})] }}),
            ]
          }}),
          {action_rows_js}
        ]
      }}),

      // 四、補充備註
      new Paragraph({{ spacing: {{ before: 240, after: 120 }}, children: [new TextRun({{ text: "四、補充備註／未決事項追蹤", bold: true, size: 32, font: "Microsoft JhengHei" }})] }}),
      {pending_js}
      {remarks_js}
    ]
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync({json.dumps(output_path)}, buffer);
  console.log("Done:", {json.dumps(output_path)});
}});
"""

# ─────────────────────────────────────────
# 傳送訊息工具
# ─────────────────────────────────────────
def send_push_message(user_id, text):
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        from linebot.v3.messaging import PushMessageRequest
        line_bot_api.push_message(
            PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=text)]
            )
        )

def send_file_to_line(user_id, file_path, meeting_info):
    """LINE 不支援直接傳 .docx，改用上傳到臨時空間後傳連結"""
    # 這裡使用 file.io 免費臨時上傳服務（7天有效）
    with open(file_path, "rb") as f:
        response = requests.post(
            "https://file.io",
            files={"file": (os.path.basename(file_path), f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            data={"expires": "7d", "autoDelete": "true"}
        )

    if response.status_code == 200:
        data = response.json()
        download_url = data.get("link", "")
        filename = os.path.basename(file_path)
        msg = f"✅ 會議記錄已完成！\n\n📄 {filename}\n\n🔗 下載連結（7天內有效）：\n{download_url}\n\n點擊連結即可下載 Word 檔案。"
    else:
        msg = "⚠️ 會議記錄已生成，但上傳時發生問題，請重試。"

    send_push_message(user_id, msg)

# ─────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
