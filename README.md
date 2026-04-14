# Automated Exam Delivery System

![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-4285F4)
![UI](https://img.shields.io/badge/ui-HTML%20%2B%20Tailwind-0EA5E9)
![Status](https://img.shields.io/badge/status-active-success)

A Google Apps Script project that automates weekly revision-paper delivery to students by email, tracks delivery logs, and provides a web dashboard for student management.

## What This Project Does

This project runs on top of Google Sheets, Google Drive, Gmail, and Apps Script:

- Sends weekly PDF papers to active students automatically.
- Reads student data from a `Students` sheet.
- Logs successes and failures to a `Logs` sheet.
- Increments each student's `NextWeek` value after successful delivery.
- Provides a browser dashboard for:
  - Viewing student and run stats
  - Searching, sorting, and paging students
  - Toggling student active/inactive status
  - Adding new students
  - Triggering paper delivery manually

Main files:

- [`Code.gs`](Code.gs): server-side Apps Script logic and spreadsheet/drive/gmail integration.
- [`dashboard.html`](dashboard.html): client-side dashboard UI and interactions.

## Why This Project Is Useful

- Removes repetitive weekly email work for instructors.
- Reduces delivery mistakes by validating folder/file availability before sending.
- Keeps an auditable record of automation activity in `Logs`.
- Gives operators a simple control panel instead of editing rows manually.
- Scales to many students with search, sort, and pagination support.

## How To Get Started

### 1) Prerequisites

- A Google account with access to:
  - Google Apps Script
  - Google Sheets
  - Google Drive
  - Gmail
- A Google Spreadsheet containing these sheets:

`Students` headers:

| Name | Email | Grade | Active | NextWeek | LastSentDate |
|---|---|---|---|---|---|

`Logs` headers:

| Timestamp | Email | Message |
|---|---|---|

- A Drive folder named `RevisionPapers` with grade subfolders (for example `Grade9`, `Grade10`, `Grade11`, `Grade12`).
- Weekly PDF files in each grade folder using this naming format:
  - `week01.pdf`
  - `week02.pdf`
  - `week03.pdf`
  - etc.

### 2) Set Up In Google Apps Script

1. Create a new Apps Script project bound to your spreadsheet (or open an existing bound script).
2. Replace/add script files with repository content:
   - Paste [`Code.gs`](Code.gs) into your `.gs` script file.
   - Paste [`dashboard.html`](dashboard.html) into an HTML file named `dashboard`.
3. Save the project.
4. Run `sendWeeklyPaper` once from the Apps Script editor to grant permissions.
5. Deploy as a web app (or test directly from script editor) for dashboard access.

### 3) Usage

Open the web app and use the dashboard to manage delivery operations.

Typical manual run flow:

1. Add or verify students in `Students`.
2. Confirm expected weekly PDFs exist in Drive folders.
3. Click **Send Papers Now** in the dashboard.
4. Review latest entries in activity logs.

Core server functions:

```javascript
// Web dashboard entry point
function doGet()

// Sends weekly paper emails to active students
function sendWeeklyPaper()

// Dashboard data provider (stats, students, logs)
function getDashboardData()

// Dashboard actions
function addStudent(student)
function updateStudentStatus(email, status)
```

## Where To Get Help

- Review the source first:
  - [`Code.gs`](Code.gs)
  - [`dashboard.html`](dashboard.html)
- Google Apps Script docs: https://developers.google.com/apps-script
- Apps Script HTML Service docs: https://developers.google.com/apps-script/guides/html
- Open an issue in this repository with:
  - your sheet structure
  - sample failing log message
  - expected vs actual behavior

## Who Maintains And Contributes

Maintainer:

- Tharin Edirisinghe

Contribution guidelines (short form):

1. Fork the repo and create a focused feature/fix branch.
2. Keep changes small and scoped (UI, automation logic, or data model).
3. Add clear reproduction steps in your PR description.
4. Include before/after behavior and any required sheet/folder assumptions.
5. Submit a pull request for review.

## Notes

- Keep student data private and avoid committing real personal data.
- This README intentionally avoids full API and troubleshooting detail to stay onboarding-focused.
