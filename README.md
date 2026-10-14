# Price Tracker Scraper

Pulls every product from an online catalogue into a clean spreadsheet, then keeps a price history so you can see what went up or down since the last run.

This demo runs against [books.toscrape.com](https://books.toscrape.com), a public practice site made for scraping. The same approach works for supplier price lists, competitor catalogues, directories and listings.

## What it does

- Walks through all 50 catalogue pages by itself (follows the "next" button)
- Collects title, price, in stock or not, star rating and link for all 1,000 books
- Removes duplicates, so a product listed twice is only counted once
- Saves the results three ways:
  - `latest_prices.csv` for any tool
  - `prices.xlsx` with a **Latest prices** tab and a **Price changes** tab (bold headers, filters, frozen top row)
  - `price_history.csv`, one row per book per day it was checked
- Optionally overwrites a Google Sheet with the latest prices
- Waits a second between pages and retries on network hiccups, so it's polite to the website

## How to run it

You need Python 3.10 or newer.

```bash
pip install -r requirements.txt
python scraper.py
```

Files land in the `output/` folder. Run it again another day and the **Price changes** tab lists every book whose price moved, with the old price, the new price and the difference.

Handy options:

| Option | What it does |
| --- | --- |
| `--max-pages 3` | Only scrape the first 3 pages (good for a quick test) |
| `--out-dir reports` | Save files somewhere else |
| `--delay 2` | Wait longer between pages |
| `--sheet "Book Prices"` | Also update a Google Sheet with that name |

### Sending results to Google Sheets (optional)

1. In Google Cloud, create a service account and download its key as `service_account.json` into this folder.
2. Share your Google Sheet with the service account's email address (as Editor).
3. Run `python scraper.py --sheet "Your Sheet Name"`.

The key file is listed in `.gitignore` so it never gets uploaded by accident.

### Running it on a schedule

To track prices daily, schedule `python scraper.py` with cron (Mac/Linux), Task Scheduler (Windows), or a GitHub Actions workflow.

## How it was tested

The `tests/` folder holds saved copies of two catalogue pages, so the tests run without internet:

```bash
python -m unittest discover tests
```

They check that every field is read correctly, that pagination is followed, that duplicates are removed, that out-of-stock items are detected, that price changes between two runs are caught, and that the Excel file has the right tabs and row count.

## How I build these

I'm an automation specialist. I use no-code tools (Zapier, Make, n8n) when they fit, and AI coding tools (Claude Code) when a job needs custom code. I then test the result myself against real data before delivering it. Scraping work is limited to public data, and I don't get around logins or bot protection.
