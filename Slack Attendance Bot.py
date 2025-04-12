import threading
import schedule
import time
from datetime import datetime, timedelta
import os
import pandas as pd
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# ✅ Generate weekly Excel filename
def get_weekly_excel_filename():
    today = datetime.now()
    start_of_week = today - timedelta(days=today.weekday())  # Monday
    end_of_week = start_of_week + timedelta(days=4)  # Friday
    start_str = start_of_week.strftime("%m%d")
    end_str = end_of_week.strftime("%m%d")
    return f"attendance_{start_str}_{end_str}.xlsx"

EXCEL_FILE = get_weekly_excel_filename()

# ✅ Slack Tokens and channel setup
SLACK_BOT_TOKEN = "input bot token"
SLACK_APP_TOKEN = "input app token"
AUTHORIZED_CHANNEL_ID = "input channel ID"

app = App(token=SLACK_BOT_TOKEN)
client = WebClient(token=SLACK_BOT_TOKEN)

# ✅ Send check-in reminder message
def send_checkin_message():
    now = datetime.now()
    weekday = now.weekday()
    if weekday < 5:
        try:
            client.chat_postMessage(
                channel=AUTHORIZED_CHANNEL_ID,
                text=":rotating_light: Time to check in! Please type `/check in` to mark your attendance. :alarm_clock:"
            )
            print(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] Check-in reminder sent")
        except SlackApiError as e:
            print(f"Failed to send message: {e.response['error']}")

# ⏰ Schedule check-in reminders
def start_scheduler():
    for minute in range(0, 60, 10):
        schedule.every().monday.at(f"09:{minute:02d}").do(send_checkin_message)
        schedule.every().tuesday.at(f"09:{minute:02d}").do(send_checkin_message)
        schedule.every().wednesday.at(f"09:{minute:02d}").do(send_checkin_message)
        schedule.every().thursday.at(f"09:{minute:02d}").do(send_checkin_message)
        schedule.every().friday.at(f"09:{minute:02d}").do(send_checkin_message)
    while True:
        schedule.run_pending()
        time.sleep(1)

# ✅ Calculate weekly work hours
def get_weekly_work_hours(name):
    total_seconds = 0
    checkin_time = None
    if not os.path.exists(EXCEL_FILE):
        return 0
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=name)
    except:
        return 0
    now = datetime.now()
    monday = now - timedelta(days=now.weekday())
    df['time'] = pd.to_datetime(df['time'])
    df_this_week = df[df['time'] >= monday.replace(hour=0, minute=0, second=0, microsecond=0)]
    for _, row in df_this_week.iterrows():
        t = row['timestamp']
        if row['status'] == 'checkin':
            checkin_time = t
        elif row['status'] == 'checkout' and checkin_time:
            start, end = checkin_time, t
            worked = (end - start).total_seconds()
            lunch_start = start.replace(hour=12, minute=0, second=0)
            lunch_end = start.replace(hour=13, minute=0, second=0)
            if start < lunch_end and end > lunch_start:
                worked -= (min(end, lunch_end) - max(start, lunch_start)).total_seconds()
            worked = min(worked, 36000)  # max 10 hours per day
            total_seconds += max(0, worked)
            checkin_time = None
    return round(total_seconds / 3600, 2)

# ✅ Get most recent status
def get_last_status(name):
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=name)
        if df.empty:
            return None
        return df.iloc[-1]['status']
    except:
        return None

# ✅ Save record
def save_record(name, status, total, remaining):
    now = datetime.now()
    timestamp = now.strftime('%Y-%m-%d %H:%M:%S')
    weekday = now.strftime('%A')
    row = pd.DataFrame([[name, timestamp, status, weekday, total, remaining]],
                       columns=["name", "timestamp", "status", "weekday", "total", "remaining"])
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            row.to_excel(writer, sheet_name=name, index=False)
    else:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            try:
                existing = pd.read_excel(EXCEL_FILE, sheet_name=name)
                updated = pd.concat([existing, row], ignore_index=True)
                writer.book.remove(writer.book[name])
                updated.to_excel(writer, sheet_name=name, index=False)
            except:
                row.to_excel(writer, sheet_name=name, index=False)
    return timestamp, weekday

# ✅ /checkin command
@app.command("/checkin")
def checkin(ack, body):
    name = body["user_name"]
    channel = body["channel_id"]
    if channel != AUTHORIZED_CHANNEL_ID:
        ack({"response_type": "ephemeral", "text": ":warning: This command is only allowed in the designated attendance channel."})
        return
    if get_last_status(name) == "checkin":
        ack({"response_type": "in_channel", "text": f":warning: {name} is already checked in. Please check out first."})
        return
    total = get_weekly_work_hours(name)
    remaining = max(0, round(40 - total, 2))
    timestamp, weekday = save_record(name, "checkin", total, remaining)
    ack({
        "response_type": "in_channel",
        "text": f":white_check_mark: {name} checked in!\n:clock3: {timestamp} ({weekday})\n:stopwatch: Total hours: {total} hrs\n:clock4: Remaining: {remaining} hrs"
    })

# ✅ /checkout command
@app.command("/checkout")
def checkout(ack, body):
    name = body["user_name"]
    channel = body["channel_id"]
    if channel != AUTHORIZED_CHANNEL_ID:
        ack({"response_type": "ephemeral", "text": ":warning: This command is only allowed in the designated attendance channel."})
        return
    if get_last_status(name) != "checkin":
        ack({"response_type": "in_channel", "text": f":warning: {name} is not currently checked in. Please check in first."})
        return
    total = get_weekly_work_hours(name)
    remaining = max(0, round(40 - total, 2))
    timestamp, weekday = save_record(name, "checkout", total, remaining)
    ack({
        "response_type": "in_channel",
        "text": f":wave: {name} checked out!\n:clock3: {timestamp} ({weekday})\n:stopwatch: Total hours: {total} hrs\n:clock4: Remaining: {remaining} hrs"
    })

# ✅ Start app and scheduler in parallel
def start_both():
    threading.Thread(target=start_scheduler, daemon=True).start()
    handler = SocketModeHandler(app, SLACK_APP_TOKEN)
    handler.start()

# ✅ Run
start_both()
