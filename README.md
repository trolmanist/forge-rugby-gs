# Forge Rugby Google Sheets Apps Script

Forge Rugby is a public Google Sheets workspace for crowdsourcing rugby player ratings during the 2026 Super Rugby Pacific season. Contributors use a custom sidebar to submit rating and weighting changes, while the sheet surfaces the live tables, rollups, and leaderboards.

This repository is the source mirror for the bound Google Apps Script that powers that spreadsheet.

## Live spreadsheet

- [Forge Rugby Ratings - SRP](https://docs.google.com/spreadsheets/d/1gMO8YiIW9S-jQBDHGzI2JLXSR7pX8_mY3TYu5v5Wmwc/edit?gid=1925124497)

## What's in this repo

- [`Code.gs`](./Code.gs): server-side Apps Script for menus, data loading, validation, submission logging, and summary rollups.
- [`Sidebar.html`](./Sidebar.html): sidebar interface for filtering players, editing values, and submitting changes.

## What this repo is for

- Transparency: the live spreadsheet uses a bound Apps Script, and this repo exposes the exact code behind that public workflow.
- Version history: GitHub provides a cleaner audit trail for changes than the Apps Script editor alone.
- Collaboration: people can inspect the logic without needing direct edit access to the spreadsheet project.

## What this repo is not

- It is not a standalone web app.
- It does not include the spreadsheet data model itself.
- The Google Sheet is the live product; this repository contains the script files used by that sheet.

## Development notes

This project currently lives as a bound Apps Script attached to the spreadsheet. If local syncing is added later via `clasp`, this repository can become the main source of truth for script changes while the spreadsheet remains the runtime host.
