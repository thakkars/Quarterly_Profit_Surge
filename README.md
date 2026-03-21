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
Go to: GitHub repo → Settings → Secrets and variables → Actions → New repository secret

| Secret Name          | Value                                      |
|----------------------|--------------------------------------------|
| SCREENER_USERNAME    | your screener.in email                     |
| SCREENER_PASSWORD    | your screener.in password                  |
| GOOGLE_CREDENTIALS   | entire contents of Google service JSON key |
| GDRIVE_FOLDER_ID     | ID from your Google Drive folder URL       |

## How to get GDRIVE_FOLDER_ID
Open your Google Drive folder in browser.
The URL looks like: https://drive.google.com/drive/folders/ABC123XYZ
The folder ID is the part after /folders/ → ABC123XYZ

## How to get GOOGLE_CREDENTIALS
1. Go to console.cloud.google.com
2. Create project → Enable "Google Drive API"
3. IAM & Admin → Service Accounts → Create
4. Keys tab → Add Key → JSON → Download
5. Open the JSON file → Copy ALL contents → paste as secret

## Running the Screener
- Go to your GitHub repo → Actions tab
- Click "Run Stock Screener" → "Run workflow" → "Run workflow"
- Takes ~10-30 mins depending on company count
- Results auto-saved to Google Drive when done
