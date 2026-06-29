# Telegram Cover Page Bot

Ei bot Telegram theke lab report ba assignment cover page generate kore DOCX file hisebe send kore.

## Features

- Menu button: `Lab Report Generator`
- Menu button: `Assignment Generator`
- Dynamic input collection:
  - University name
  - Course code
  - Course title
  - Assignment/report no
  - Student name
  - Student ID
  - Program
  - Batch/year/semester
  - Teacher name
  - Teacher designation
  - Teacher department
  - University logo upload
- Render webhook support
- Local polling support

## Render setup

1. BotFather theke Telegram bot token nin.
2. Ei project GitHub-e upload korun.
3. Render-e new Web Service create korun.
4. Environment variables add korun:
   - `BOT_TOKEN`: BotFather token
   - `WEBHOOK_URL`: Render service URL, example `https://cover-page-telegram-bot.onrender.com`
5. Deploy korun.

## Local run

`.env` file use korle manually environment variable set korte hobe, ba terminal theke:

```powershell
$env:BOT_TOKEN="YOUR_BOT_TOKEN"
python bot.py
```

Local run-e `WEBHOOK_URL` na dile bot polling mode-e cholbe.
