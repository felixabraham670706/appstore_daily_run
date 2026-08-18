#!/usr/bin/env python3
"""
App Store Ratings Tracker — Multi-Bank Edition
================================================
Fetches today's rating data for a Google Play app and an Apple App Store
app for EACH bank listed in the BANKS config below, then:

  1. For every bank, appends ONE ROW to a fixed "<CODE> Google Play" tab
     and ONE ROW to a fixed "<CODE> Apple App Store" tab — every field as
     a column, "Date" as the first column so you can filter/sort by day.
     Re-running the same day updates that day's row instead of adding a
     duplicate.
  2. For every bank, updates a running "<CODE> Summary" tab that keeps one
     row per day showing:
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
     applies that month), also writes one row into each bank's
     "<CODE> Monthly Summary" tab with that month's total new ratings and
     average rating per platform (same weighted-average trick as the daily
     figures, just applied across the whole month).
  4. Emails the resulting workbook (ALL banks, one file) as an attachment
     to whichever addresses are listed in RECIPIENT_EMAILS (see the email
     config section below — nothing here is hardcoded; it all comes from
     environment variables).

BANKS COVERED (9 total, 4 tabs each = 36 tabs in one workbook):
    ENBD (Emirates NBD), EI (Emirates Islamic), Mashreq, Wio, ADCB,
    ADCB NEW (ADCB's separate redesigned app), ADIB, FAB, RAK (RAKBANK)

Tabs are named "<CODE> Summary", "<CODE> Google Play", "<CODE> Apple App
Store" and "<CODE> Monthly Summary" for each bank, e.g. "ENBD Summary",
"EI Google Play", "FAB Apple App Store", etc.

IMPORTANT — please double-check the app IDs before relying on this in
production. The Google Play package names and Apple App Store numeric IDs
for EI, Mashreq, Wio, ADCB, ADIB, FAB and RAK below were identified via web
search (Aug 2026) by matching each bank's official store listing name/
developer to their own website. ADCB, FAB and RAK in particular each
publish several similarly-named apps (personal banking vs business/forex/
securities apps), so it's worth opening each store URL in the comments
below once to confirm it's the exact app you want tracked. If a bank ever
relaunches under a new package/app ID, update its entry in BANKS — nothing
else in the script needs to change.

NOTE ON THE TWO ADCB ENTRIES: "ADCB" tracks their original/legacy personal
banking app (Google Play com.adcb.bank + Apple id 547172388); "ADCB NEW"
tracks their separate, newly redesigned "ADCB." app (Google Play
com.adcb.nexgen + Apple id 6755109454). These are two distinct apps, so
each gets its own 4 tabs.

HEADS UP — HISTORY DISCONTINUITY: the "ADCB" entry's google_app_id was
corrected on 2026-08-18 from com.adcb.nexgen to com.adcb.bank (the actual
Play Store counterpart of Apple id 547172388). If you had already run this
tracker with the old (mismatched) setting, the "ADCB Google Play" and
"ADCB Summary" tabs will show a one-time jump in Total Ratings / Daily New
Ratings on the date of this fix, since the totals switch from one app to a
different one. That's expected and only affects that single transition
day — everything from this point forward is tracking the correct app.

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
# CONFIG — one entry per bank. Add/remove/edit banks here only; every
# other function below works off this list, so nothing else needs to
# change to add an 8th bank later.
#
#   code          -> short code used as the tab-name prefix, e.g. "ENBD"
#   display_name  -> full name used in the log and the email body
#   google_app_id -> Google Play package id (from the play.google.com URL)
#   google_country-> Play Store country code to fetch from
#   apple_app_id  -> numeric Apple App Store id (from the apps.apple.com URL)
#   apple_country -> App Store country code to fetch from
# ----------------------------------------------------------------------
BANKS = [
    {
        "code": "ENBD",
        "display_name": "Emirates NBD",
        "google_app_id": "com.emiratesnbd.android",
        "google_country": "ae",
        "apple_app_id": "1497518128",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.emiratesislamic.android
        # https://apps.apple.com/ae/app/ei/id1499264261
        "code": "EI",
        "display_name": "Emirates Islamic",
        "google_app_id": "com.emiratesislamic.android",
        "google_country": "ae",
        "apple_app_id": "1499264261",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.vipera.ts.starter.MashreqAE
        # https://apps.apple.com/ae/app/mashreq-uae-digital-banking/id378549193
        "code": "MASHREQ",
        "display_name": "Mashreq",
        "google_app_id": "com.vipera.ts.starter.MashreqAE",
        "google_country": "ae",
        "apple_app_id": "378549193",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=io.wio.retail
        # https://apps.apple.com/ae/app/wio-personal/id1658472726
        "code": "WIO",
        "display_name": "Wio Bank",
        "google_app_id": "io.wio.retail",
        "google_country": "ae",
        "apple_app_id": "1658472726",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.adcb.bank
        # https://apps.apple.com/ae/app/adcb/id547172388
        # ADCB's original/legacy personal banking app ("ADCB", plain name,
        # no period — confirmed by you on 2026-08-18). ADCB also publishes
        # ADCB Hayyak, ADCB ProCash, ADCB Nomo, ADCB Business, etc. — this
        # entry is meant to be their main personal banking app.
        "code": "ADCB",
        "display_name": "Abu Dhabi Commercial Bank",
        "google_app_id": "com.adcb.bank",
        "google_country": "ae",
        "apple_app_id": "547172388",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.adcb.nexgen
        # https://apps.apple.com/ae/app/adcb/id6755109454
        # ADCB's separate, newly redesigned "ADCB." app (with a period —
        # AI-powered, US equities/crypto trading, etc.) — kept as its own
        # bank entry rather than replacing the ADCB entry above, per your
        # request to track both.
        "code": "ADCB NEW",
        "display_name": "ADCB (New App)",
        "google_app_id": "com.adcb.nexgen",
        "google_country": "ae",
        "apple_app_id": "6755109454",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.adib.mobile
        # https://apps.apple.com/ae/app/adib-mobile-banking/id1128180440
        "code": "ADIB",
        "display_name": "Abu Dhabi Islamic Bank",
        "google_app_id": "com.adib.mobile",
        "google_country": "ae",
        "apple_app_id": "1128180440",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.fab.personalbanking
        # https://apps.apple.com/ae/app/fab-mobile-banking/id1383237548
        # NOTE: FAB also publishes FAB Business, FAB Securities, FABeAccess,
        # etc. — this entry is meant to be their main personal banking app;
        # please verify against your own FAB app store links.
        "code": "FAB",
        "display_name": "First Abu Dhabi Bank",
        "google_app_id": "com.fab.personalbanking",
        "google_country": "ae",
        "apple_app_id": "1383237548",
        "apple_country": "ae",
    },
    {
        # https://play.google.com/store/apps/details?id=com.rak
        # https://apps.apple.com/ae/app/rakbank/id427758991
        # NOTE: RAKBANK also publishes "RAKBANK Business" (package
        # com.rakcorp) — this entry is their main personal banking app;
        # please verify against your own RAKBANK app store links.
        "code": "RAK",
        "display_name": "RAKBANK",
        "google_app_id": "com.rak",
        "google_country": "ae",
        "apple_app_id": "427758991",
        "apple_country": "ae",
    },
]

# Folder this script lives in — the workbook & log are kept next to it
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "app_ratings_history.xlsx")
LOG_PATH = os.path.join(BASE_DIR, "ratings_tracker.log")

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
# SHEET NAMING — every bank gets 4 tabs, prefixed with its short code so
# all 7 banks live side by side in one workbook without colliding.
# ----------------------------------------------------------------------
def bank_sheet_names(code: str) -> dict:
    return {
        "summary": f"{code} Summary",
        "google": f"{code} Google Play",
        "apple": f"{code} Apple App Store",
        "monthly": f"{code} Monthly Summary",
    }


def migrate_legacy_enbd_sheets(wb: Workbook):
    """
    The very first version of this tracker (ENBD-only) used unprefixed
    tab names: "Summary", "Google Play", "Apple App Store", "Monthly
    Summary". If a workbook created by that old version is loaded here,
    rename those tabs in place to "ENBD Summary" etc. so the existing
    history is kept instead of starting a second, disconnected ENBD tab
    set. This only ever runs once — after the rename, the legacy names
    are gone, so later runs are a no-op.
    """
    legacy_to_new = {
        "Summary": "ENBD Summary",
        "Google Play": "ENBD Google Play",
        "Apple App Store": "ENBD Apple App Store",
        "Monthly Summary": "ENBD Monthly Summary",
    }
    for old_name, new_name in legacy_to_new.items():
        if old_name in wb.sheetnames and new_name not in wb.sheetnames:
            wb[old_name].title = new_name
            log_and_print(f"Migrated legacy tab '{old_name}' -> '{new_name}'")


def reorder_sheets(wb: Workbook, banks: list):
    """
    Rebuilds tab order so all 4 tabs for a bank sit together, and banks
    appear in the same order as the BANKS config list — regardless of the
    order sheets happened to be created/recreated in during this run.
    Any sheet that doesn't match the expected naming (shouldn't normally
    happen) is left at the end rather than dropped.
    """
    desired_order = []
    for bank in banks:
        names = bank_sheet_names(bank["code"])
        desired_order.extend([names["summary"], names["google"], names["apple"], names["monthly"]])

    ordered_existing = [name for name in desired_order if name in wb.sheetnames]
    leftover = [name for name in wb.sheetnames if name not in ordered_existing]
    wb._sheets = [wb[name] for name in ordered_existing + leftover]


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
    One FIXED tab per platform per bank (e.g. "ENBD Google Play"). Column A
    is "Date", every other column is one field from `data`. Each day adds
    one new row at the bottom, so you can filter/sort the whole tab by
    Date.

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


def update_summary_sheet(wb: Workbook, sheet_name: str, day_str: str,
                          g_total: float, g_avg: float,
                          a_total: float, a_avg: float):
    """
    Rebuilds one bank's Summary sheet from scratch every run: keeps every
    previous day's numbers, adds/replaces today's row, recomputes the two
    "daily" columns per platform, and redraws both charts.

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
    if sheet_name in wb.sheetnames:
        ws_old = wb[sheet_name]
        for row in ws_old.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            existing_rows.append(list(row))
        del wb[sheet_name]

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

    ws = wb.create_sheet(sheet_name)
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
    bank_label = ws.title.replace(" Summary", "")

    dates_ref = Reference(ws, min_col=1, min_row=2, max_row=last_row)

    chart1 = LineChart()
    chart1.title = f"{bank_label} — Daily New Ratings — Google Play vs Apple App Store"
    chart1.style = 2
    chart1.y_axis.title = "New ratings that day"
    chart1.x_axis.title = "Date"
    chart1.height, chart1.width = 9, 22
    chart1.add_data(Reference(ws, min_col=4, min_row=1, max_row=last_row), titles_from_data=True)
    chart1.add_data(Reference(ws, min_col=8, min_row=1, max_row=last_row), titles_from_data=True)
    chart1.set_categories(dates_ref)
    ws.add_chart(chart1, "K2")

    chart2 = LineChart()
    chart2.title = f"{bank_label} — Daily Average Rating — Google Play vs Apple App Store"
    chart2.style = 10
    chart2.y_axis.title = "Avg rating that day"
    chart2.x_axis.title = "Date"
    chart2.height, chart2.width = 9, 22
    chart2.add_data(Reference(ws, min_col=5, min_row=1, max_row=last_row), titles_from_data=True)
    chart2.add_data(Reference(ws, min_col=9, min_row=1, max_row=last_row), titles_from_data=True)
    chart2.set_categories(dates_ref)
    ws.add_chart(chart2, "K20")


def update_monthly_summary(wb: Workbook, summary_sheet_name: str, monthly_sheet_name: str, day_str: str):
    """
    Only meant to be called on the last calendar day of the month, for one
    bank at a time. Reads the day-by-day numbers already sitting in that
    bank's Summary sheet, works out this month's total NEW ratings and
    this month's average rating (same weighted-average trick as the daily
    calculation in update_summary_sheet, just applied across a stretch of
    days instead of one day), and writes one row into that bank's
    "<CODE> Monthly Summary" tab — keyed by "YYYY-MM" so re-running on the
    same month-end day updates that row instead of duplicating it.

    Baseline used for the delta — always the FIRST tracked day of THIS
    calendar month (ideally the 1st), never a day borrowed from the
    previous month:
      - Normally that's the row dated "<month>-01" — e.g. August 1 when
        today is August 31 — so the month's new ratings and month's
        average rating are worked out purely from that day's and today's
        Total Ratings / Average Rating, the same weighted-average trick
        used for the daily figures in update_summary_sheet.
      - If tracking only started partway through the month (no row on the
        1st), it falls back to the EARLIEST row you do have this month, so
        the numbers still reflect "new ratings since tracking began this
        month" rather than sitting blank. That fallback is flagged in the
        "Notes" column so it's obvious the figure covers a partial month,
        not the full one.
      - If today is the ONLY row that exists this month (tracking started
        and hit month-end on the very same day), there's genuinely nothing
        to diff against, and the figures stay blank.
    """
    if summary_sheet_name not in wb.sheetnames:
        return None

    ws_summary = wb[summary_sheet_name]
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

    # baseline = the earliest row THIS MONTH — ideally the 1st of the month
    this_month_rows = [r for r in rows if r[0] >= first_day_this_month]
    baseline = this_month_rows[0] if this_month_rows else None

    is_partial_month = False
    if baseline is not None and baseline[0] == today_row[0]:
        # today is the only row tracked this month — nothing to diff against
        baseline = None
    elif baseline is not None and baseline[0] != first_day_this_month:
        is_partial_month = True

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

    notes = f"Partial month — first tracked day was {baseline[0]}" if is_partial_month else ""

    header = [
        "Month",
        "GPlay Total Ratings (EOM)", "GPlay New Ratings This Month", "GPlay Avg Rating This Month",
        "Apple Total Ratings (EOM)", "Apple New Ratings This Month", "Apple Avg Rating This Month",
        "Notes",
    ]

    if monthly_sheet_name in wb.sheetnames:
        ws_m = wb[monthly_sheet_name]
    else:
        ws_m = wb.create_sheet(monthly_sheet_name)
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
# PER-BANK PROCESSING
# ----------------------------------------------------------------------
def process_bank(wb: Workbook, bank: dict, today, today_str: str) -> dict:
    """
    Fetches and records one bank's data. Unlike a hard crash, a failure
    for ONE bank (a scraper hiccup, a renamed app id, a store outage) is
    caught and logged here so the other 6 banks still get processed, the
    workbook still gets saved, and the email still goes out — the failed
    bank is just flagged in the log and in the email body instead of
    silently vanishing or blocking everyone else.
    """
    code = bank["code"]
    display_name = bank["display_name"]
    result = {"code": code, "display_name": display_name, "ok": False, "error": None}

    try:
        google_data = fetch_google_data(bank["google_app_id"], bank["google_country"])
        log_and_print(
            f"[{code}] Google Play OK — Total Ratings={google_data['Total Ratings']}, "
            f"Avg={google_data['Average Rating']}"
        )
    except Exception:
        log_and_print(f"[{code}] FAILED fetching Google Play data:\n" + traceback.format_exc(), "error")
        result["error"] = "Google Play fetch failed (see log)"
        return result

    try:
        apple_data = fetch_apple_data(bank["apple_app_id"], bank["apple_country"])
        log_and_print(
            f"[{code}] Apple App Store OK — Total Ratings={apple_data['Total Ratings']}, "
            f"Avg={apple_data['Average Rating']}"
        )
    except Exception:
        log_and_print(f"[{code}] FAILED fetching Apple data:\n" + traceback.format_exc(), "error")
        result["error"] = "Apple App Store fetch failed (see log)"
        return result

    try:
        names = bank_sheet_names(code)
        append_or_update_platform_row(wb, names["google"], today_str, google_data)
        append_or_update_platform_row(wb, names["apple"], today_str, apple_data)
        update_summary_sheet(
            wb, names["summary"], today_str,
            g_total=float(google_data["Total Ratings"] or 0),
            g_avg=float(google_data["Average Rating"] or 0),
            a_total=float(apple_data["Total Ratings"] or 0),
            a_avg=float(apple_data["Average Rating"] or 0),
        )

        last_day_of_month = calendar.monthrange(today.year, today.month)[1]
        if today.day == last_day_of_month:
            update_monthly_summary(wb, names["summary"], names["monthly"], today_str)
            log_and_print(f"[{code}] Month-end detected ({today_str}) — updated {names['monthly']} tab.")
    except Exception:
        log_and_print(f"[{code}] FAILED writing Excel data:\n" + traceback.format_exc(), "error")
        result["error"] = "Excel update failed (see log)"
        return result

    result.update({
        "ok": True,
        "google_total": google_data["Total Ratings"],
        "google_avg": google_data["Average Rating"],
        "apple_total": apple_data["Total Ratings"],
        "apple_avg": apple_data["Average Rating"],
    })
    return result


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    today = date.today()
    today_str = today.isoformat()
    log_and_print(f"=== Ratings tracker run: {today_str} — {len(BANKS)} bank(s) ===")

    wb = get_or_create_workbook(EXCEL_PATH)
    migrate_legacy_enbd_sheets(wb)

    results = [process_bank(wb, bank, today, today_str) for bank in BANKS]

    reorder_sheets(wb, BANKS)

    try:
        wb.save(EXCEL_PATH)
        log_and_print(f"Saved workbook: {EXCEL_PATH}")
    except Exception:
        log_and_print("FAILED writing Excel workbook:\n" + traceback.format_exc(), "error")
        sys.exit(1)

    body_lines = [f"Attached is the latest ratings workbook as of {today_str}.", ""]
    any_failure = False
    for r in results:
        if r["ok"]:
            body_lines.append(
                f"{r['display_name']} ({r['code']}) — Google Play: {r['google_total']} ratings "
                f"(avg {r['google_avg']}) | Apple App Store: {r['apple_total']} ratings "
                f"(avg {r['apple_avg']})"
            )
        else:
            any_failure = True
            body_lines.append(f"{r['display_name']} ({r['code']}) — FAILED: {r['error']}")

    send_email_with_attachment(
        attachment_path=EXCEL_PATH,
        subject=f"Bank App Store Ratings Update — {today_str}",
        body="\n".join(body_lines),
    )

    if any_failure:
        log_and_print("=== Done (with failures — see FAILED lines above) ===\n", "warning")
        sys.exit(1)

    log_and_print("=== Done ===\n")


if __name__ == "__main__":
    main()