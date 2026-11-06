# Stock movers -> Telegram (runs free on GitHub Actions)

Every weekday ~15:45 IST it reads the stocks in `data/*.xlsx`, fetches closing prices
(Yahoo Finance, NSE/BSE) and sends you the ones that went **up today**, biggest first.
Skips holidays automatically.

## One-time setup
1. Telegram: message **@BotFather** -> `/newbot` -> copy the **bot token**.
   Send any message to your new bot, then open
   `https://api.telegram.org/bot<TOKEN>/getUpdates` and copy `chat.id`.
2. Create a **private** GitHub repo and push the contents of this folder
   (`stock_alerts/`) to it (`.github/` must be at the repo root).
3. Repo -> Settings -> Secrets and variables -> Actions:
   - Secrets: `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`
   - (optional) Variable `MIN_GAIN_PCT` e.g. `3` for only +3% or more. Default: every stock that is up.
4. Actions tab -> "Daily stock movers alert" -> Run workflow (enter `1` in force to test on a closed day).

## Updating the lists
Replace the files in `data/` (same sheet and column names) and push.

## Wrong / missing ticker
Tickers are auto-matched by company name on first run (see `map`/`MISS` lines in the
Actions log). Fix or add any in `symbols.csv` (`name,symbol` - NSE symbol, or BSE code for numeric).

## Local test (needs Python 3.10+)
    pip install -r requirements.txt
    set DRY_RUN=1 && set FORCE=1 && python stock_alert.py
