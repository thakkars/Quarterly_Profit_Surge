# Stock Screener — GitHub Automation & Dashboard

100% hosted and automated on GitHub (No Cloudflare dependency).

## Live Web Dashboards (GitHub Pages)
- **Main Screener Dashboard:** [https://thakkars.github.io/Quarterly_Profit_Surge/](https://thakkars.github.io/Quarterly_Profit_Surge/)
- **NSE Results Calendar:** [https://thakkars.github.io/Quarterly_Profit_Surge/results_calendar.html](https://thakkars.github.io/Quarterly_Profit_Surge/results_calendar.html)

---

## How to Enable GitHub Pages
1. Go to your GitHub repository: **`thakkars/Quarterly_Profit_Surge`**
2. Click **Settings** → **Pages** (in the left sidebar).
3. Under **Build and deployment** → **Source**, choose: **Deploy from a branch**.
4. Select branch: **`main`** and folder: **`/ (root)`**.
5. Click **Save**.
6. Within a minute, your dashboard will be live at: `https://thakkars.github.io/Quarterly_Profit_Surge/`.

---

## Required GitHub Secrets
Go to: GitHub repo → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

| Secret Name          | Description / Value                                      |
|----------------------|----------------------------------------------------------|
| `SCREENER_USERNAME`  | Your Screener.in account email                           |
| `SCREENER_PASSWORD`  | Your Screener.in account password                        |
| `TELEGRAM_BOT_TOKEN` | Bot token from `@BotFather` (e.g. `123456:ABC-DEF...`)    |
| `TELEGRAM_CHAT_ID`   | Telegram User ID, Group ID, or comma-separated list of IDs (e.g. `123456789, -1001234567890`) |

*(Cloudflare R2 secrets are no longer required!)*

---

## Workflow Permissions (Required for Auto-Commit)
For GitHub Actions to automatically commit and push updated data back to your repo:
1. Go to **Settings** → **Actions** → **General**.
2. Scroll to **Workflow permissions**.
3. Choose **Read and write permissions**.
4. Check **"Allow GitHub Actions to create and approve pull requests"**.
5. Click **Save**.

---

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

---

## Running the Screener
- Go to your GitHub repo → **Actions** tab.
- Click **"Run Stock Screener"** → **Run workflow**.
- When complete, an HTML summary with matched breakout stocks will be sent directly to your Telegram, and the updated `rawdata.json` is committed automatically.
- You can also test Telegram notifications instantly anytime via the **"Test Telegram Alert"** workflow.
