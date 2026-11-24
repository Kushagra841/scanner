# Packing Scan Log

A web page for packing staff. A USB barcode scanner types the code into the page, and each scan is saved with the packer's name, set/camera number, and the exact date and time down to the millisecond. Every scan is sent to one central Google Sheet.

Live page: https://kushagra841.github.io/scanner/

## Files

| File | What it is |
|---|---|
| `index.html` | The page packers use (hosted by GitHub Pages) |
| `scan_server.gs` | The Google Apps Script inside the Google Sheet that receives and stores the scans |

## What it does

- Saves each scan on the PC first, then sends it to the Google Sheet. If the internet drops, scans wait and are sent later, 25 at a time.
- Warns when a barcode was already scanned before, by anyone, including scans in old archive tabs.
- Admin password needed to download scans or clear the on-screen list.
- Download one day, or all days, as an Excel-ready CSV file.
- On the 1st of every month, scans older than last month move to tabs like `Archive 2026-09`, so the main `Scans` tab stays fast.

## Updating the Google script

The real admin password is **not** in this repo. `scan_server.gs` has a placeholder.

1. Open the Google Sheet → **Extensions → Apps Script**.
2. Copy your current `const ADMIN_KEY = "...";` line.
3. Replace all the code with `scan_server.gs` from here, then put your `ADMIN_KEY` line back.
4. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.** (Not "New deployment": that changes the web address the page sends to.)
5. First time only: run `setupMonthlyArchive` once to switch on the monthly archive.

`SECRET` must be the same in `scan_server.gs` and `index.html`.
