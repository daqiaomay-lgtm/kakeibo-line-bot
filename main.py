import base64, hmac, hashlib, os, requests, pytz, gspread
from datetime import datetime, timedelta
from dateutil import parser as dtparser
from typing import List, Dict, Any
from fastapi import FastAPI, Request, Header, HTTPException

app = FastAPI()
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
SHEET_ID = os.getenv("SHEET_ID", "")
TZ = os.getenv("TZ", "Asia/Tokyo")
CREDS = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "credentials.json")

def get_sheet():
    gc = gspread.service_account(filename=CREDS)
    return gc.open_by_key(SHEET_ID).worksheet("log")

def sum_by_period(rows: List[List[str]], period: str, tzname: str) -> int:
    tz = pytz.timezone(tzname); now = datetime.now(tz)
    if period == "today":
        start = tz.localize(datetime(now.year, now.month, now.day)); end = start + timedelta(days=1)
    elif period == "week":
        start = tz.localize(datetime(now.year, now.month, now.day)) - timedelta(days=now.weekday()); end = start + timedelta(days=7)
    else:
        start = tz.localize(datetime(now.year, now.month, 1))
        end = tz.localize(datetime(now.year + (now.month==12), (now.month%12)+1, 1))
    total = 0
    for r in rows[1:]:
        if len(r) < 2: continue
        ds, amt = r[0], r[1]
        if not ds or not amt: continue
        try:
            d = dtparser.parse(ds); d = tz.localize(d) if d.tzinfo is None else d.astimezone(tz)
            a = float(str(amt).replace(",", ""))
        except Exception:
            continue
        if start <= d < end: total += a
    return int(total)

def verify_signature(secret: str, body: bytes, sig: str) -> bool:
    mac = hmac.new(secret.encode("utf-8"), body, hashlib.sha256).digest()
    return hmac.compare_digest(base64.b64encode(mac).decode("utf-8"), sig)

def line_reply(token: str, text: str):
    r = requests.post("https://api.line.me/v2/bot/message/reply",
        headers={"Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}", "Content-Type": "application/json"},
        json={"replyToken": token, "messages":[{"type":"text","text":text}]}, timeout=10)
    r.raise_for_status()

@app.get("/health")
def health(): return {"status":"ok"}

@app.post("/webhook")
async def webhook(request: Request, x_line_signature: str = Header(None)):
    if not (LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN and SHEET_ID):
        raise HTTPException(status_code=500, detail="env not set")
    body_bytes = await request.body()
    if not x_line_signature or not verify_signature(LINE_CHANNEL_SECRET, body_bytes, x_line_signature):
        raise HTTPException(status_code=401, detail="Invalid signature")

    events: List[Dict[str, Any]] = (await request.json()).get("events", [])
    for ev in events:
        if ev.get("type")!="message" or ev.get("message",{}).get("type")!="text": continue
        text = ev["message"]["text"].strip()
        if text in ("今月","今週","今日"):
            period = "month" if text=="今月" else ("week" if text=="今週" else "today")
            try:
                rows = get_sheet().get_all_values()
                total = sum_by_period(rows, period, TZ)
                label = {"month":"今月","week":"今週","today":"今日"}[period]
                reply = f"{label}の累計：{total:,} 円"
            except Exception as e:
                reply = f"エラー：{e}\nSHEET_ID/共有/credentials.json を確認してね。"
        else:
            reply = "使い方：\n「今日」「今週」「今月」を送ってください。"
        try: line_reply(ev["replyToken"], reply)
        except Exception as e: print("Reply error:", e)
    return {"status":"ok"}
