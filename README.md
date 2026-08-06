# App Store Ratings Tracker

Automatically fetches daily rating data for a Google Play app and an Apple
App Store app, and keeps a growing Excel workbook (`app_ratings_history.xlsx`) with:

- **A `Google Play` tab** — one row per day, every field as a column, `Date`
  as the first column so you can filter/sort by day
- **An `Apple App Store` tab** — same idea, one row per day
- **A `Summary` tab** — one row per day with:
  - Total ratings & average rating (as shown on the store), both platforms
  - **Daily new ratings** = today's total − yesterday's total
  - **Daily average rating** = an estimate of the rating people gave *that day only*
    (the stores only ever show a cumulative average — see the comment above
    `update_summary_sheet()` in the script for the exact math)
  - Two auto-updating line charts: daily new ratings, and daily average rating,
    both showing Google Play vs Apple

Runs once a day via cron and always appends — nothing is ever overwritten
except the current day's data if you re-run it twice in a day.

## Files

| File | Purpose |
|---|---|
| `app_ratings_tracker.py` | The script that does everything |
| `requirements.txt` | Python packages it needs |
| `run_daily.sh` | Optional cron wrapper that also auto-pushes the workbook to GitHub |
| `app_ratings_history.xlsx` | Created automatically the first time you run it |

## 1. One-time setup

You need Python 3.9+ and git installed on the machine that will run this
daily (your own laptop/PC left on, or a small cloud server — cron only runs
while the machine is on).

```bash
# check you have python3 and git
python3 --version
git --version
```

Clone this repo (after you've pushed it to GitHub) or just create a folder
and put the files in it:

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
```

Create a virtual environment and install dependencies:

```bash
python3 -m venv venv
source venv/bin/activate          # on Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Configure your app IDs

Open `app_ratings_tracker.py` and edit the top of the file:

```python
GOOGLE_APP_ID = "com.emiratesnbd.android"   # your Play Store package name
GOOGLE_COUNTRY = "ae"

APPLE_APP_ID = "1497518128"                 # your App Store numeric ID
APPLE_COUNTRY = "ae"
```

These already match the app from your notebook — only change them if you
ever need to track a different app.

## 3. Test it manually before automating anything

```bash
source venv/bin/activate
python3 app_ratings_tracker.py
```

You should see log lines print to your terminal, and two new files appear:
`app_ratings_history.xlsx` and `ratings_tracker.log`. Open the Excel file
and check:
- a `Google Play` tab with one row (today's date + all fields)
- an `Apple App Store` tab with one row
- a `Summary` tab with one row and two (mostly empty, until day 2) charts

Run it again the next day (or fake it by changing your system clock — not
recommended) and you'll see a second row appear in `Google Play`, `Apple App
Store`, and `Summary`, with the daily delta columns in `Summary` now filled
in. Re-running the script twice on the *same* day just updates that day's
row everywhere instead of adding a duplicate.

## 4. Automate it with cron (runs daily at 8:00 AM)

Open your crontab editor:

```bash
crontab -e
```

Add this line (edit the paths to match where you actually put things —
cron doesn't know about your virtual environment, so we call the venv's
python directly):

```cron
0 8 * * * /full/path/to/your-repo/venv/bin/python3 /full/path/to/your-repo/app_ratings_tracker.py >> /full/path/to/your-repo/cron_output.log 2>&1
```

Find your full path with `pwd` while inside the repo folder. Save and exit
— cron will now run the script every day at 8:00 AM as long as the machine
is on at that time.

**Check it worked:** the next day after 8 AM, look at `ratings_tracker.log`
inside the repo — it logs every run, success or failure, with a timestamp.

## 5. (Optional) Auto-push the growing Excel file to GitHub

If you'd like the workbook itself backed up to GitHub after every run
(so you have version history and can view it from anywhere), use
`run_daily.sh` instead of calling the `.py` file directly from cron:

1. Edit `REPO_DIR` inside `run_daily.sh` to your repo's full path.
2. Make it executable: `chmod +x run_daily.sh`
3. Make sure `git push` won't ask for a password interactively — set up
   an SSH key or a
   [personal access token](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens)
   with git configured to use it, and test `git push` manually once first.
4. Point cron at the wrapper instead:

```cron
0 8 * * * /full/path/to/your-repo/run_daily.sh >> /full/path/to/your-repo/cron_output.log 2>&1
```

## 6. Push the code itself to GitHub (do this once)

```bash
cd your-repo
git init                     # only if you haven't already
git add app_ratings_tracker.py requirements.txt run_daily.sh .gitignore README.md
git commit -m "Initial ratings tracker"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

The `.gitignore` already excludes your `venv/` folder so you don't
accidentally commit thousands of library files.

## Troubleshooting

- **Nothing happens at 8 AM** → check `crontab -l` shows your line exactly,
  and that the machine was actually on/awake at that time.
- **Script fails immediately when run manually** → read the error printed
  in the terminal; the most common cause is forgetting to
  `source venv/bin/activate` first, or a typo in an app ID.
- **`Average Rating` / `Total Ratings` come back empty** → double check the
  app IDs are correct for the store you're pulling from (Play Store uses
  the package name like `com.company.app`; App Store uses the numeric ID
  from the app's store URL).
- **A "Daily Average Rating" value looks impossible (e.g. way outside
  1–5)** → this usually means the store recalculated/adjusted its historical
  totals between your two readings, not a bug in the script — the
  back-calculation assumes both days' figures are internally consistent.
