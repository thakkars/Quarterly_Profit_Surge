# Stock Screener — Cloud Setup Guide

## Project Structure
```
stock-screener/
├── R01_70__Operating_Profit.py   ← main script (cloud-ready)
├── download_from_drive.py         ← downloads files from Drive before run
├── requirements.txt               ← Python dependencies
├── data/                          ← Excel files saved here (git-ignored)
└── .github/
    └── workflows/
        └── run_screener.yml       ← GitHub Actions automation
```

## GitHub Secrets to Add
Go to: GitHub repo → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Secret Name            | Description / Value                                      |
|------------------------|----------------------------------------------------------|
| `SCREENER_USERNAME`    | Your Screener.in account email                           |
| `SCREENER_PASSWORD`    | Your Screener.in account password                        |
| `R2_ACCESS_KEY_ID`     | Cloudflare R2 Access Key ID                              |
| `R2_SECRET_ACCESS_KEY` | Cloudflare R2 Secret Access Key                          |
| `R2_ENDPOINT_URL`      | Cloudflare R2 Endpoint URL                               |
| `R2_BUCKET_NAME`       | Cloudflare R2 Bucket Name (default: `quarterly-profit-surge`) |
| `TELEGRAM_BOT_TOKEN`   | Bot token from `@BotFather` (e.g. `123456:ABC-DEF...`)    |
| `TELEGRAM_CHAT_ID`     | Telegram User ID, Group ID, or comma-separated list of IDs (e.g. `123456789, -1001234567890`) |

## Setting Up Telegram Alerts (Personal & Group)
1. Open Telegram and search for `@BotFather`.
2. Send `/newbot`, name your bot, and copy the **API Token** (`TELEGRAM_BOT_TOKEN`).
3. Start a chat with your new bot and click **Start** (`/start`).
4. To find your **Personal Chat ID**:
   - Search for `@userinfobot` or `@GetIDsBot` on Telegram and send `/start`.
   - Copy your numeric `Id` (e.g. `123456789`).
5. To send to a **Telegram Group**:
   - Add your bot to your group and promote it to **Admin**.
   - Send any test message in your group (e.g., `hello`).
   - Open this URL in your browser: `https://api.telegram.org/bot<YOUR_BOT_TOKEN>/getUpdates`
   - Look for `"chat":{"id": -100xxxxxxxxxx...}` in the response. Copy that negative number (including `-` or `-100`).
6. Set `TELEGRAM_CHAT_ID` in your GitHub Secrets:
   - For both Personal + Group: `123456789, -100xxxxxxxxxx` (separated by comma).


## Running the Screener
- Go to your GitHub repo → **Actions** tab.
- Click **"Run Stock Screener"** → **Run workflow**.
- When complete, an HTML summary with matched breakout stocks will be sent directly to your Telegram.
- You can also test Telegram notifications instantly anytime via the **"Test Telegram Alert"** workflow.

