import os
import json
import base64
import httpx
import anthropic
from fastapi import FastAPI, Request, BackgroundTasks, HTTPException
from fastapi.responses import JSONResponse
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi,
    ReplyMessageRequest, TextMessage
)
from linebot.v3.webhooks import MessageEvent, ImageMessageContent
from linebot.v3.exceptions import InvalidSignatureError
from datetime import datetime

app = FastAPI()

LINE_CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
GAS_WEB_APP_URL = os.environ["GAS_WEB_APP_URL"]

handler = WebhookHandler(LINE_CHANNEL_SECRET)
line_config = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
claude_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


@app.get("/")
def root():
    return {"status": "Yanmar Nameplate OCR Service is running"}


@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):
    signature = request.headers.get("X-Line-Signature", "")
    body = await request.body()
    try:
        handler.handle(body.decode("utf-8"), signature)
    except InvalidSignatureError:
        raise HTTPException(status_code=400, detail="Invalid signature")
    return JSONResponse(content={"status": "ok"})


@handler.add(MessageEvent, message=ImageMessageContent)
def handle_image(event):
    with ApiClient(line_config) as api_client:
        line_bot_api = MessagingApi(api_client)
        message_content = line_bot_api.get_message_content(event.message.id)
        image_data = base64.b64encode(message_content.read()).decode("utf-8")

    result = read_nameplate(image_data)

    if result["success"]:
        save_to_sheet(result["data"])
        d = result["data"]
        reply_text = (
            f"\u2705 อ่านป้ายสำเร็จ!\n"
            f"\U0001f4cb Model: {d['model']}\n"
            f"\U0001f522 Serial No.: {d['serial_no']}\n"
            f"\u2699\ufe0f Engine Displacement: {d['engine_displacement']}\n"
            f"\U0001f4c5 {d['timestamp']}\n\n"
            f"บันทึกลง Google Sheets แล้วครับ \u2713"
        )
    else:
        reply_text = f"\u274c ไม่สามารถอ่านป้ายได้\nสาเหตุ: {result['error']}"

    with ApiClient(line_config) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=reply_text)]
            )
        )


def read_nameplate(image_base64: str) -> dict:
    try:
        response = claude_client.messages.create(
            model="claude-opus-4-5",
            max_tokens=500,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": image_base64,
                        },
                    },
                    {
                        "type": "text",
                        "text": "Read the nameplate and reply in JSON only:\n{\"model\": \"...\", \"serial_no\": \"...\", \"engine_displacement\": \"...\"}\nUse null if a value cannot be read."
                    }
                ],
            }],
        )
        raw = response.content[0].text.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        data = json.loads(raw)
        data["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return {"success": True, "data": data}
    except Exception as e:
        return {"success": False, "error": str(e)}


def save_to_sheet(data: dict):
    try:
        with httpx.Client(timeout=15) as client:
            client.post(GAS_WEB_APP_URL, json=data)
    except Exception as e:
        print(f"[GAS Error] {e}")
