#!/bin/bash
# ----------------------------------------------------------------------
# Wrapper for cron. Activates the venv, runs the tracker, and (optionally)
# commits + pushes the updated Excel file back to your GitHub repo.
#
# Point crontab at THIS script instead of the .py file directly if you
# want the workbook auto-backed-up to GitHub after every run.
# If you don't want that, just point cron straight at app_ratings_tracker.py
# and ignore/delete this file.
# ----------------------------------------------------------------------

# EDIT THIS to the folder where you cloned/placed the repo
REPO_DIR="/home/youruser/app-ratings-tracker"

cd "$REPO_DIR" || exit 1

# activate the virtual environment
source venv/bin/activate

# run the tracker
python3 app_ratings_tracker.py

# commit & push the updated workbook (only if something actually changed)
git add app_ratings_history.xlsx ratings_tracker.log
if ! git diff --cached --quiet; then
    git commit -m "Daily ratings update: $(date +'%Y-%m-%d')"
    git push origin main
fi
