# Slack Attendance Bot

![그림2](https://github.com/user-attachments/assets/4269ce60-46a4-44af-96c9-86eeda7dfb23)

This is a Slack-based attendance bot that allows users to check in and out using `/출근` and `/퇴근` commands. It also automatically sends reminder messages every weekday morning at 10-minute intervals from 09:00 to 09:50.

## ✅ Features

- Slack slash commands: `/checkin`, `/checkout`
- Weekly Excel file storage: `attendance_MMDD_MMDD.xlsx`
- Each user has their own Excel sheet
- Lunch break (12:00–13:00) is excluded from work hours
- Maximum 10 hours per day counted
- Weekly reset (starts every Monday)
- Only works in a designated Slack channel
- Automated reminders every weekday at 09:00–09:50 every 10 minutes

## 🛠 Requirements

Install required Python packages:

```bash
pip install -r requirements.txt![Uploading 그림1.png…]()

```

## ★ Running

```bash
python3 attendance_bot.py
