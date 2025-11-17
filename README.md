# Leave-Schedule-Tracker-Google-Sheets-Apps-Script-
A lightweight leave-scheduling system built on Google Sheets and Google Apps Script. Designed for small teams who need a simple way to track leave days.

Features

Auto-generated monthly sheets for an entire year
Per-employee protected ranges (each employee can only edit their own column)
Prevents leave on closed/holiday days
Prevents exceeding leave limits
Summary sheet with total & remaining leave days
Daily automated Drive backups (keeps last N backups)
Admin-only backup log

Setup Instructions
1. Create a New Google Spreadsheet
2. Open the Apps Script Editor: Extensions → Apps Script, copy the entire project code into Code.gs.
3. Configure settings at the top of the code (Back up folder in Google Drive needed for back up functionality)
4. Run generateMonthlySheets() once
5. Add Triggers in Apps Scripts for the onEdit() function - triggers on edit, and for the backupSpreadsheetDaily() function - triggers whenever backup is needed to take place
