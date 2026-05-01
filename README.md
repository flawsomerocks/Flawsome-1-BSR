# Daily BSR Updater

Automatically fetches the current Amazon Best Seller Rank for each ASIN in your Google Sheet and writes it to the **Current BSR** column (E) every day via GitHub Actions.

---

## How it works

1. The GitHub Actions workflow runs daily at 06:00 UTC.
2. It reads every ASIN from **column A** of your *BSR Tracking* sheet.
3. For each ASIN it visits `https://www.amazon.com/dp/<ASIN>` and parses the BSR.
4. It writes the result back to **column E** of the same row.

---

## One-time setup

### 1. Create a Google Cloud service account

1. Go to [Google Cloud Console](https://console.cloud.google.com/) → **IAM & Admin** → **Service Accounts**.
2. Click **Create Service Account**, give it a name (e.g. `bsr-updater`).
3. On the service account page, go to **Keys** → **Add Key** → **Create new key** → **JSON**. Download the file.

### 2. Share your Google Sheet with the service account

1. Open the downloaded JSON file and copy the `client_email` value.
2. In Google Sheets, click **Share** and paste that email with **Editor** access.

### 3. Enable the Google Sheets API

In Google Cloud Console → **APIs & Services** → **Library**, search for and enable:
- **Google Sheets API**
- **Google Drive API**

### 4. Add GitHub Secrets & Variables

In your GitHub repo: **Settings → Secrets and variables → Actions**

| Type | Name | Value |
|------|------|-------|
| Secret | `GOOGLE_SHEETS_CREDENTIALS` | Paste the entire contents of the service account JSON file |
| Secret | `SPREADSHEET_ID` | The ID from your sheet URL: `.../spreadsheets/d/<ID>/edit` |
| Variable (optional) | `SHEET_NAME` | Tab name — defaults to `BSR Tracking` |
| Variable (optional) | `AMAZON_MARKETPLACE` | `com`, `co.uk`, `de`, etc. — defaults to `com` |

### 5. Trigger a test run

Go to **Actions** → **Daily BSR Update** → **Run workflow** to verify everything works before waiting for the scheduled run.

---

## Adjusting the schedule

Edit the `cron` expression in `.github/workflows/daily_bsr_update.yml`:

```yaml
- cron: "0 6 * * *"   # 06:00 UTC every day
```

Use [crontab.guru](https://crontab.guru/) to generate your preferred schedule.

---

## Sheet layout expected

| A | B | C | D | E |
|---|---|---|---|---|
| ASIN | Link | Best Seller Ranking | Best Seller Ranking | **Current BSR** ← updated here |

Rows without an ASIN in column A are skipped automatically.

---

## Running locally

```bash
pip install -r requirements.txt

export GOOGLE_SHEETS_CREDENTIALS='<paste JSON>'
export SPREADSHEET_ID='<your sheet ID>'

python update_bsr.py
```
