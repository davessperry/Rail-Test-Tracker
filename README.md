# Rail Test Tracker

A mobile-friendly web app for tracking rail test mileage across subdivisions (Cascade, Thompson, Shuswap, Page, Mission, Westminster) — remaining miles by track type, Sidings/Xovers checklists, daily XML report import, and a running log of what's been tested.

**Live app:** https://claude.ai/artifact/D4Qav1bP5prPX3Hs5e7UNi

## What's here

- `index.html` — the entire app: a single self-contained HTML/CSS/JS file.
- `apple-touch-icon.png` — the icon used when the app is added to an iPhone home screen.

## Built-in track list

Edit master data > "Add a subdivision from the Master tracks list" lets anyone search 178 subdivisions and add one (its tracks, sidings and crossovers) with a tap. The list is a snapshot of the company Master tracks list dated August 17th, 2026 (track names and mileposts only), embedded in `index.html` as `BUILTIN_TRACK_LIST`. A newer spreadsheet can be loaded from the same card.

## How this repo is used

This repo is a backup and change history for the app's source code. The actual running app lives on Claude's Artifact platform (the "Live app" link above), which also provides the app's data storage (remaining miles, test history, checklist state). Opening `index.html` directly from this repo (e.g. via GitHub Pages) will **not** have working data storage, since that storage is a feature of the Artifact platform, not something built into the file itself.

Changes are made through a Claude conversation and pushed here afterward, so this history reflects what changed and when — independent of any single chat session.

## Daily report add-on (optional)

`libreoffice/DailyReport.bas` is a LibreOffice Calc macro. `libreoffice/DailyReport.oxt` packages it as an extension (rebuild with `python3 libreoffice/build_oxt.py`); installing it adds a **Daily Report** menu to Calc. After tapping **Copy Remaining Miles** or **Copy Miles Tested Today** in the app, choose the matching item: it reads the clipboard, finds the labels in column B, and refills the cells to their right (adding rows if needed). The install steps and the .oxt download are in the app under **Spreadsheet macro**.
