import os, io, datetime as dt
import pandas as pd
import pytz
from fastapi import FastAPI, Request
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from linebot.exceptions import InvalidSignatureError

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# ========= 環境変数 =========
TZ = os.getenv("TZ", "Asia/Tokyo")
JST = pytz.timezone(TZ)
SHEET_ID = os.getenv("SHEET_ID")
EXCEL_FILE_ID = os.getenv("EXCEL_FILE_ID")  # 既存ID。空なら初回作成
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")

# ========= FastAPI / LINE =========
app = FastAPI()
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# ========= Google 認証 =========
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
credentials = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
gc = gspread.authorize(credentials)
drive = build("drive", "v3", credentials=credentials)

def get_sheet():
    sh = gc.open_by_key(SHEET_ID)
    return sh.worksheet("log")  # 入力用シート

# ========= Excel ユーティリティ =========
def download_excel_to_df(file_id: str) -> pd.DataFrame:
    request = drive.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    try:
        df = pd.read_excel(fh, sheet_name="log")
    except Exception:
        df = pd.DataFrame(columns=["日付", "金額", "使った人"])
    if df.empty:
        df = pd.DataFrame(columns=["日付", "金額", "使った人"])
    return df

def create_or_update_excel_append(rows: list) -> str:
    """rows: [[日付, 金額, 使った人], ...] をExcelに追記（無ければ新規作成）"""
    global EXCEL_FILE_ID
    add_df = pd.DataFrame(rows, columns=["日付", "金額", "使った人"])

    if EXCEL_FILE_ID:
        base_df = download_excel_to_df(EXCEL_FILE_ID)
        merged = pd.concat([base_df, add_df], ignore_index=True)
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as w:
            merged.to_excel(w, index=False, sheet_name="log")
        bio.seek(0)
        media = MediaIoBaseUpload(
            bio,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        )
        drive.files().update(fileId=EXCEL_FILE_ID, media_body=media).execute()
        return EXCEL_FILE_ID
    else:
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="openpyxl") as w:
            add_df.to_excel(w, index=False, sheet_name="log")
        bio.seek(0)
        file_metadata = {"name": "kakeibo_log.xlsx", "mimeType": "application/vnd.google-apps.file"}
        media = MediaIoBaseUpload(
            bio,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        )
        created = drive.files().create(body=file_metadata, media_body=media, fields="id").execute()
        EXCEL_FILE_ID = created["id"]
        return EXCEL_FILE_ID

def clear_sheet(ws):
    ws.clear()

# ========= 集計（Excel基準） =========
def sum_by_range_from_excel(when: str) -> int:
    if not EXCEL_FILE_ID:
        return 0
    df = download_excel_to_df(EXCEL_FILE_ID)
    if df.empty:
        return 0

    df["日付"] = pd.to_datetime(df["日付"], errors="coerce").dt.date
    df["金額"] = pd.to_numeric(df["金額"], errors="coerce").fillna(0).astype(int)

    today = dt.datetime.now(JST).date()
    if when == "今日":
        flt = df["日付"] == today
    elif when == "今週":
        start = today - dt.timedelta(days=today.weekday())  # 月曜始まり
        end = start + dt.timedelta(days=6)
        flt = (df["日付"] >= start) & (df["日付"] <= end)
    else:  # 今月
        start = today.replace(day=1)
        next_month_1 = (start.replace(year=start.year + 1, month=1, day=1)
                        if start.month == 12 else start.replace(month=start.month + 1, day=1))
        end = next_month_1 - dt.timedelta(days=1)
        flt = (df["日付"] >= start) & (df["日付"] <= end)
    return int(df.loc[flt, "金額"].sum())

# ========= データ移送（保存） =========
def move_sheet_to_excel_and_clear():
    ws = get_sheet()
    values = ws.get_all_values()
    if not values:
        return 0
    rows = values[1:] if len(values) > 1 else []
    if not rows:
        return 0

    cleaned = []
    for r in rows:
        if len(r) >= 2 and r[0] and r[1]:
            cleaned.append([r[0], r[1], r[2] if len(r) > 2 else ""])

    if not cleaned:
        return 0

    create_or_update_excel_append(cleaned)
    clear_sheet(ws)
    return len(cleaned)

# ========= ルーティング =========
@app.post("/webhook")
async def webhook(request: Request):
    signature = request.headers.get('X-Line-Signature')
    body = await request.body()
    try:
        handler.handle(body.decode('utf-8'), signature)
    except InvalidSignatureError:
        return "Invalid signature"
    return "OK"

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event: MessageEvent):
    text = event.message.text.strip()

    if text in ["今日", "今週", "今月"]:
        total = sum_by_range_from_excel(text)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"{text}の累計：{total:,}円")
        )
        return

    if text in ["保存", "save"]:
        n = move_sheet_to_excel_and_clear()
        msg = f"{n}件をExcelへ移動し、シートを初期化しました。"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        return

    help_text = "使い方：\n・保存 → シート→Excelへ退避\n・今日/今週/今月 → Excelから集計して返信"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=help_text))
