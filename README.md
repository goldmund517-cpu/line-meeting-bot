# LINE 會議記錄 Bot

自動將 LINE 語音訊息轉換為正式 Word 會議記錄。

## 流程

```
LINE 語音 → Gemini 逐字稿+語者辨識 → 確認語者 → GPT 生成記錄 → Word 檔下載連結
```

## 環境變數（在 Render 填入）

| 變數名稱 | 說明 |
|---------|------|
| `LINE_CHANNEL_SECRET` | LINE Channel Secret |
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Channel Access Token |
| `GEMINI_API_KEY` | Google Gemini API Key |
| `OPENAI_API_KEY` | OpenAI API Key |

## Render 設定

- **Build Command**: `pip install -r requirements.txt && npm install`
- **Start Command**: `gunicorn main:app --bind 0.0.0.0:$PORT`
- **Runtime**: Python 3.11

## LINE Webhook 設定

Render 部署完成後，將以下網址填入 LINE Developers：
```
https://你的render網址.onrender.com/callback
```
