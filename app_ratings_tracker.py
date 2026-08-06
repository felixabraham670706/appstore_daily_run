#!/usr/bin/env python3
"""
App Store Ratings Tracker
==========================
Fetches today's rating data for a Google Play app and an Apple App Store
app, then:

  1. Appends ONE ROW to a fixed "Google Play" tab and ONE ROW to a fixed
     "Apple App Store" tab — every field as a column, "Date" as the first
     column so you can filter/sort by day. Re-running the same day updates
     that day's row instead of adding a duplicate.
  2. Updates a running "Summary" tab that keeps one row per day showing:
       - total ratings & average rating (as shown on the store) for each
         platform
       - how many NEW ratings came in that specific day (today's total
         minus yesterday's total)
       - an estimated AVERAGE RATING for just that day's new ratings
         (see the note in `update_summary_sheet` for how this is derived)
     ...and (re)draws two line charts on it:
       - Daily new ratings, Google Play vs Apple
       - Daily average rating, Google Play vs Apple
  3. On the LAST CALENDAR DAY of the month (28th/29th/30th/31st, whichever
     applies that month), also writes one row into a "Monthly Summary" tab
     with that month's total new ratings and average rating per platform
     (same weighted-average trick as the daily figures, just applied
     across the whole month).
  4. Emails the resulting workbook as an attachment to whichever addresses
     are listed in RECIPIENT_EMAILS (see the email config section below —
     nothing here is hardcoded; it all comes from environment variables).

Designed to be run once a day (e.g. via a scheduled GitHub Actions
workflow, or cron) via:
    python3 app_ratings_tracker.py
"""

import os
import sys
import logging
import traceback
import smtplib
import calendar
from datetime import date
from email.message import EmailMessage

import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

# Loads variables from a local ".env" file if one exists next to this
# script (handy for testing on your own PC). In GitHub Actions there is
# no .env file — the same variable names are injected as real environment
# variables from repo Secrets instead, so this call is simply a no-op there.
load_dotenv()

# ----------------------------------------------------------------------
# CONFIG — edit these for your app(s)
# ----------------------------------------------------------------------
GOOGLE_APP_ID = "com.emiratesnbd.android"
GOOGLE_COUNTRY = "ae"

APPLE_APP_ID = "1497518128"
APPLE_COUNTRY = "ae"

# Folder this script lives in — the workbook & log are kept next to it
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "app_ratings_history.xlsx")
LOG_PATH = os.path.join(BASE_DIR, "ratings_tracker.log")

SUMMARY_SHEET = "Summary"
GOOGLE_SHEET = "Google Play"
APPLE_SHEET = "Apple App Store"
MONTHLY_SHEET = "Monthly Summary"

# ----------------------------------------------------------------------
# EMAIL CONFIG — every value below comes from an environment variable.
# NEVER hardcode an email address's password directly in this file —
# this file gets pushed to git. Set these instead:
#
#   Locally (testing on your PC):
#       create a file named ".env" next to this script (never commit it —
#       it's already in .gitignore) containing:
#           SENDER_EMAIL=youraccount@gmail.com
#           SENDER_PASSWORD=your_16_char_app_password
#           RECIPIENT_EMAILS=personal@gmail.com,you@yourcompany.com
#           SMTP_SERVER=smtp.gmail.com
#           SMTP_PORT=587
#
#   In GitHub Actions (scheduled runs):
#       add the same names as repo Secrets — Settings → Secrets and
#       variables → Actions → New repository secret — and reference them
#       in the workflow yml's `env:` block (see README).
#
# Note: you only need ONE authenticated sender account. Both your personal
# and office addresses just go in RECIPIENT_EMAILS, comma-separated — you
# do NOT need your office mailbox's own SMTP credentials.
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# LOGGING — so a scheduled run (which you never watch live) leaves a trail
# ----------------------------------------------------------------------
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


def log_and_print(msg, level="info"):
    print(msg)
    getattr(logger, level)(msg)


# ----------------------------------------------------------------------
# FETCHERS  (same fields as your original notebook)
# ----------------------------------------------------------------------
def fetch_google_data(app_id: str, country: str) -> dict:
    from google_play_scraper import app as gplay_app  # imported here so the

    # rest of the script can be tested even before this package is installed
    result = gplay_app(app_id, lang="en", country=country)

    return {
        "App ID": result.get("appId"),
        "App Name": result.get("title"),
        "Summary": result.get("summary"),
        "Description": result.get("description"),
        "Developer": result.get("developer"),
        "Developer ID": result.get("developerId"),
        "Developer Website": result.get("developerWebsite"),
        "Developer Email": result.get("developerEmail"),
        "Category": result.get("genre"),
        "Category ID": result.get("genreId"),
        "Current Version": result.get("version"),
        "Released": result.get("released"),
        "Last Updated": result.get("updated"),
        "Size": result.get("size"),
        "Installs": result.get("installs"),
        "Minimum Installs": result.get("minInstalls"),
        "Average Rating": result.get("score"),
        "Displayed Rating": result.get("scoreText"),
        "Total Ratings": result.get("ratings"),
        "Total Reviews": result.get("reviews"),
        "Content Rating": result.get("contentRating"),
        "Free": result.get("free"),
        "Price": result.get("price"),
        "Currency": result.get("currency"),
        "Contains Ads": result.get("containsAds"),
        "In App Purchases": result.get("offersIAP"),
        "Privacy Policy": result.get("privacyPolicy"),
        "Icon": result.get("icon"),
        "Header Image": result.get("headerImage"),
        "URL": result.get("url"),
    }


def fetch_apple_data(app_id: str, country: str) -> dict:
    url = f"https://itunes.apple.com/lookup?id={app_id}&country={country}"
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    payload = response.json()

    if not payload.get("results"):
        raise ValueError(f"No results returned by iTunes lookup for app id {app_id}")

    data = payload["results"][0]

    return {
        "App Name": data.get("trackName"),
        "Developer": data.get("sellerName"),
        "Average Rating": data.get("averageUserRating"),
        "Total Ratings": data.get("userRatingCount"),
        "Current Version Rating": data.get("averageUserRatingForCurrentVersion"),
        "Current Version Ratings Count": data.get("userRatingCountForCurrentVersion"),
        "Current Version": data.get("version"),
        "Minimum OS": data.get("minimumOsVersion"),
        "Price": data.get("formattedPrice"),
        "Release Date": data.get("releaseDate"),
        "Last Updated": data.get("currentVersionReleaseDate"),
        "Genre": data.get("primaryGenreName"),
        "Description": data.get("description"),
    }


# ----------------------------------------------------------------------
# EXCEL HELPERS
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=12)


def get_or_create_workbook(path: str) -> Workbook:
    if os.path.exists(path):
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet
    return wb


def append_or_update_platform_row(wb: Workbook, sheet_name: str, day_str: str, data: dict):
    """
    One FIXED tab per platform (e.g. "Google Play"). Column A is "Date",
    every other column is one field from `data`. Each day adds one new
    row at the bottom, so you can filter/sort the whole tab by Date.

    If a row for `day_str` already exists (you re-ran the script the same
    day), that row is updated in place instead of adding a duplicate.
    """
    header = ["Date"] + list(data.keys())

    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(header)
        for c in range(1, len(header) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 14
        for i in range(2, len(header) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 24
    else:
        ws = wb[sheet_name]

    # find an existing row for this date (column A), if any
    row_idx = None
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == day_str:
            row_idx = r
            break

    values = [day_str] + list(data.values())
    if row_idx:
        for c, v in enumerate(values, start=1):
            ws.cell(row=row_idx, column=c, value=v)
    else:
        ws.append(values)

    return ws


def update_summary_sheet(wb: Workbook, day_str: str,
                          g_total: float, g_avg: float,
                          a_total: float, a_avg: float):
    """
    Rebuilds the Summary sheet from scratch every run: keeps every previous
    day's numbers, adds/replaces today's row, recomputes the two "daily"
    columns per platform, and redraws both charts.

    How "daily new ratings" is calculated:
        today's Total Ratings  -  yesterday's Total Ratings

    How "daily average rating" (the rating people gave THAT day only) is
    calculated:
        The store only ever shows a cumulative average across all ratings
        ever left, never a per-day figure. We can still recover it with a
        weighted-average trick: if you know the cumulative average and
        count both today and yesterday, the average of just today's new
        ratings is
            (today_avg * today_total - yesterday_avg * yesterday_total)
            / (today_total - yesterday_total)
        This is left blank for a day with zero new ratings (can't divide
        by zero) and for the very first day ever recorded (no "yesterday"
        to compare against).
    """
    header = [
        "Date",
        "GPlay Total Ratings", "GPlay Avg Rating",
        "GPlay Daily New Ratings", "GPlay Daily Avg Rating",
        "Apple Total Ratings", "Apple Avg Rating",
        "Apple Daily New Ratings", "Apple Daily Avg Rating",
    ]

    existing_rows = []
    if SUMMARY_SHEET in wb.sheetnames:
        ws_old = wb[SUMMARY_SHEET]
        for row in ws_old.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            existing_rows.append(list(row))
        del wb[SUMMARY_SHEET]

    # keep only the "raw" numbers per day; every delta column is
    # recomputed fresh below so edits/re-runs never compound errors
    base_rows = {r[0]: (r[1], r[2], r[5], r[6]) for r in existing_rows}
    base_rows[day_str] = (g_total, g_avg, a_total, a_avg)

    final_rows = []
    prev_g_total = prev_g_avg = prev_a_total = prev_a_avg = None
    for d in sorted(base_rows.keys()):
        gt, ga, at, aa = base_rows[d]

        g_new = (gt - prev_g_total) if prev_g_total is not None else None
        a_new = (at - prev_a_total) if prev_a_total is not None else None

        g_day_avg = None
        if g_new and g_new > 0:
            g_day_avg = round((ga * gt - prev_g_avg * prev_g_total) / g_new, 3)

        a_day_avg = None
        if a_new and a_new > 0:
            a_day_avg = round((aa * at - prev_a_avg * prev_a_total) / a_new, 3)

        final_rows.append([d, gt, ga, g_new, g_day_avg, at, aa, a_new, a_day_avg])
        prev_g_total, prev_g_avg = gt, ga
        prev_a_total, prev_a_avg = at, aa

    ws = wb.create_sheet(SUMMARY_SHEET, 0)
    ws.append(header)
    for c in range(1, len(header) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for r in final_rows:
        ws.append(r)

    for col, width in zip("ABCDEFGHI", [12, 16, 14, 18, 18, 16, 14, 18, 18]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"

    if final_rows:
        _add_charts(ws, len(final_rows))

    return ws


def _add_charts(ws, n_rows):
    last_row = n_rows + 1  # +1 to account for the header row

    dates_ref = Reference(ws, min_col=1, min_row=2, max_row=last_row)

    chart1 = LineChart()
    chart1.title = "Daily New Ratings — Google Play vs Apple App Store"
    chart1.style = 2
    chart1.y_axis.title = "New ratings that day"
    chart1.x_axis.title = "Date"
    chart1.height, chart1.width = 9, 22
    chart1.add_data(Reference(ws, min_col=4, min_row=1, max_row=last_row), titles_from_data=True)
    chart1.add_data(Reference(ws, min_col=8, min_row=1, max_row=last_row), titles_from_data=True)
    chart1.set_categories(dates_ref)
    ws.add_chart(chart1, "K2")

    chart2 = LineChart()
    chart2.title = "Daily Average Rating — Google Play vs Apple App Store"
    chart2.style = 10
    chart2.y_axis.title = "Avg rating that day"
    chart2.x_axis.title = "Date"
    chart2.height, chart2.width = 9, 22
    chart2.add_data(Reference(ws, min_col=5, min_row=1, max_row=last_row), titles_from_data=True)
    chart2.add_data(Reference(ws, min_col=9, min_row=1, max_row=last_row), titles_from_data=True)
    chart2.set_categories(dates_ref)
    ws.add_chart(chart2, "K20")


def update_monthly_summary(wb: Workbook, day_str: str):
    """
    Only meant to be called on the last calendar day of the month. Reads
    the day-by-day numbers already sitting in the Summary sheet, works out
    this month's total NEW ratings and this month's average rating (same
    weighted-average trick as the daily calculation in update_summary_sheet,
    just applied across a stretch of days instead of one day), and writes
    one row into a "Monthly Summary" tab — keyed by "YYYY-MM" so re-running
    on the same month-end day updates that row instead of duplicating it.

    Baseline used for the delta:
      - Ideally, the last row from BEFORE this month started (e.g. July 31
        if today is August 31) — this gives a true full-calendar-month figure.
      - If that doesn't exist (tracking only started partway through this
        month, or this is the very first month ever tracked), it falls back
        to the EARLIEST row you do have this month, so the numbers still
        reflect "new ratings since tracking began" rather than sitting
        blank. That fallback is flagged in the "Notes" column so it's
        obvious the figure covers a partial month, not the full one.
      - If today is the ONLY row that exists (tracking started and hit
        month-end on the very same day), there's genuinely nothing to
        diff against, and the figures stay blank.
    """
    if SUMMARY_SHEET not in wb.sheetnames:
        return None

    ws_summary = wb[SUMMARY_SHEET]
    rows = []
    for row in ws_summary.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append(row)  # Date, GTotal, GAvg, GNew, GDayAvg, ATotal, AAvg, ANew, ADayAvg
    rows.sort(key=lambda r: r[0])

    today_row = next((r for r in rows if r[0] == day_str), None)
    if today_row is None:
        return None

    year, month = int(day_str[:4]), int(day_str[5:7])
    month_prefix = f"{year:04d}-{month:02d}"
    first_day_this_month = f"{month_prefix}-01"

    # preferred baseline = the most recent row strictly before this month started
    baseline = None
    for r in rows:
        if r[0] < first_day_this_month:
            baseline = r
        else:
            break

    is_partial_month = False
    if baseline is None:
        # no full-month baseline available — fall back to the earliest row
        # we DO have this month (as long as it isn't today's own row)
        this_month_rows = [r for r in rows if r[0] >= first_day_this_month]
        if len(this_month_rows) >= 2:
            baseline = this_month_rows[0]
            is_partial_month = True
        # else: today is the only row that exists — nothing to diff against

    g_today_total, g_today_avg = today_row[1], today_row[2]
    a_today_total, a_today_avg = today_row[5], today_row[6]

    if baseline:
        g_base_total, g_base_avg = baseline[1], baseline[2]
        a_base_total, a_base_avg = baseline[5], baseline[6]

        g_month_new = g_today_total - g_base_total
        a_month_new = a_today_total - a_base_total

        g_month_avg = (
            round((g_today_avg * g_today_total - g_base_avg * g_base_total) / g_month_new, 3)
            if g_month_new > 0 else None
        )
        a_month_avg = (
            round((a_today_avg * a_today_total - a_base_avg * a_base_total) / a_month_new, 3)
            if a_month_new > 0 else None
        )
    else:
        g_month_new = a_month_new = None
        g_month_avg = a_month_avg = None

    notes = f"Partial month — data from {baseline[0]}" if is_partial_month else ""

    header = [
        "Month",
        "GPlay Total Ratings (EOM)", "GPlay New Ratings This Month", "GPlay Avg Rating This Month",
        "Apple Total Ratings (EOM)", "Apple New Ratings This Month", "Apple Avg Rating This Month",
        "Notes",
    ]

    if MONTHLY_SHEET in wb.sheetnames:
        ws_m = wb[MONTHLY_SHEET]
    else:
        ws_m = wb.create_sheet(MONTHLY_SHEET)
        ws_m.append(header)
        for c in range(1, len(header) + 1):
            cell = ws_m.cell(row=1, column=c)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        ws_m.freeze_panes = "A2"
        for col, width in zip("ABCDEFGH", [12, 20, 24, 24, 20, 24, 24, 32]):
            ws_m.column_dimensions[col].width = width

    row_idx = None
    for r in range(2, ws_m.max_row + 1):
        if ws_m.cell(row=r, column=1).value == month_prefix:
            row_idx = r
            break

    values = [month_prefix, g_today_total, g_month_new, g_month_avg,
              a_today_total, a_month_new, a_month_avg, notes]
    if row_idx:
        for c, v in enumerate(values, start=1):
            ws_m.cell(row=row_idx, column=c, value=v)
    else:
        ws_m.append(values)

    return ws_m


# ----------------------------------------------------------------------
# EMAIL
# ----------------------------------------------------------------------
def send_email_with_attachment(attachment_path: str, subject: str, body: str):
    """
    Emails the current workbook to whoever is listed in RECIPIENT_EMAILS.
    All settings come from environment variables (see the EMAIL CONFIG
    comment near the top of this file). If they aren't set, this logs a
    warning and simply skips emailing — it never crashes the whole run,
    since the workbook has already been saved successfully by this point.
    """
    sender = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_PASSWORD")
    recipients_raw = os.environ.get("RECIPIENT_EMAILS")
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))

    if not sender or not password or not recipients_raw:
        log_and_print(
            "Email not sent — SENDER_EMAIL / SENDER_PASSWORD / RECIPIENT_EMAILS "
            "are not set. Skipping the email step (the workbook itself was "
            "still saved normally).",
            "warning",
        )
        return

    recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(attachment_path),
        )

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender, password)
            server.send_message(msg)
        log_and_print(f"Email sent to: {', '.join(recipients)}")
    except Exception:
        log_and_print("FAILED sending email:\n" + traceback.format_exc(), "error")


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    today = date.today()
    today_str = today.isoformat()
    log_and_print(f"=== Ratings tracker run: {today_str} ===")

    try:
        google_data = fetch_google_data(GOOGLE_APP_ID, GOOGLE_COUNTRY)
        log_and_print(
            f"Google Play OK — Total Ratings={google_data['Total Ratings']}, "
            f"Avg={google_data['Average Rating']}"
        )
    except Exception:
        log_and_print("FAILED fetching Google Play data:\n" + traceback.format_exc(), "error")
        sys.exit(1)

    try:
        apple_data = fetch_apple_data(APPLE_APP_ID, APPLE_COUNTRY)
        log_and_print(
            f"Apple App Store OK — Total Ratings={apple_data['Total Ratings']}, "
            f"Avg={apple_data['Average Rating']}"
        )
    except Exception:
        log_and_print("FAILED fetching Apple data:\n" + traceback.format_exc(), "error")
        sys.exit(1)

    try:
        wb = get_or_create_workbook(EXCEL_PATH)
        append_or_update_platform_row(wb, GOOGLE_SHEET, today_str, google_data)
        append_or_update_platform_row(wb, APPLE_SHEET, today_str, apple_data)
        update_summary_sheet(
            wb, today_str,
            g_total=float(google_data["Total Ratings"] or 0),
            g_avg=float(google_data["Average Rating"] or 0),
            a_total=float(apple_data["Total Ratings"] or 0),
            a_avg=float(apple_data["Average Rating"] or 0),
        )

        last_day_of_month = calendar.monthrange(today.year, today.month)[1]
        if today.day == last_day_of_month:
            update_monthly_summary(wb, today_str)
            log_and_print(f"Month-end detected ({today_str}) — updated Monthly Summary tab.")

        wb.save(EXCEL_PATH)
        log_and_print(f"Saved workbook: {EXCEL_PATH}")
    except Exception:
        log_and_print("FAILED writing Excel workbook:\n" + traceback.format_exc(), "error")
        sys.exit(1)

    send_email_with_attachment(
        attachment_path=EXCEL_PATH,
        subject=f"App Store Ratings Update — {today_str}",
        body=(
            f"Attached is the latest ratings workbook as of {today_str}.\n\n"
            f"Google Play — Total Ratings: {google_data['Total Ratings']}, "
            f"Average Rating: {google_data['Average Rating']}\n"
            f"Apple App Store — Total Ratings: {apple_data['Total Ratings']}, "
            f"Average Rating: {apple_data['Average Rating']}"
        ),
    )

    log_and_print("=== Done ===\n")


if __name__ == "__main__":
    main()