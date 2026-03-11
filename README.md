# Forge Rugby - Google Sheets Apps Script

The Forge Rugby ratings spreadsheet is a public Google Sheet used to crowdsource rugby player ratings for the 2026 Super Rugby Pacific season. Contributors use a custom sidebar to submit changes, while the sheet surfaces live ratings, weights, and leaderboards.

This folder contains the source for the bound Google Apps Script that powers that workflow:

- [`Code.gs`](./Code.gs): server-side Apps Script for menus, data loading, submission logging, and rollups.
- [`Sidebar.html`](./Sidebar.html): sidebar UI for submitting player rating and weighting changes.

Live spreadsheet:

- [Forge Rugby Ratings - SRP](https://docs.google.com/spreadsheets/d/1gMO8YiIW9S-jQBDHGzI2JLXSR7pX8_mY3TYu5v5Wmwc/edit?gid=1925124497)

This source is published here for transparency. The live spreadsheet uses a bound Apps Script, and because that script is currently unverified by Google, this repo provides the exact code that powers the public ratings workspace.

The spreadsheet is the live product; this folder is a source mirror for the Apps Script files used in that sheet.
