# Installation & Setup Guide

This Google Apps Script project automates weekly task distribution for Tim Hortons managers using Google Sheets.

### Prerequisites
- A Google account
- Google Sheet with the required tabs (Settings, Announcements, Logs, Leaders Board, Self-Evaluations, etc.) — make a copy of your existing one
- Basic familiarity with Google Sheets

### Step-by-Step Setup

1. **Prepare your Google Sheet**
   - Open your Tim Hortons task spreadsheet (or make a copy if sharing a template)
   - Ensure it has at least these sheets:
     - Settings (managers & emails)
     - Announcements
     - Logs (will be auto-created)
     - Leaders Board
     - Self-Evaluations
     - Calibration & Receiving Logs
     - Weekly Reports - Section Assignments
     - CLUSTER - Training Videos
     - Supervisor Task Assignments
     - Deadlines Submissions

2. **Open Apps Script editor**
   - In the Google Sheet: **Extensions → Apps Script**

3. **Add the code**
   - Delete any default code in `Code.gs`
   - Copy-paste the entire script from the `src/` folder files in this repository (or combine into one file if preferred)
   - **Recommended structure** (create separate files via + button in editor):
     - `main.gs` — onOpen(), menus, preview/send functions
     - `setup.gs` — sheet creation functions
     - `email-builder.gs` — sendEmails() and HTML generation
     - `data-collectors.gs` — collectTasks() logic
     - `utils.gs` — helpers like logAction, getManagersAndEmails

4. **Save and authorize**
   - Click the disk icon (Save project)
   - Run any function (e.g. `onOpen()`) → you will be asked to authorize permissions (Sheets, Gmail, etc.)
   - Approve all (this script needs access to read/write Sheets and send emails)

5. **Reload the Sheet**
   - Close and reopen the spreadsheet tab
   - A new **Tasks** menu should appear in the top bar

6. **Test it**
   - Go to **Tasks → Send Test Emails** → enter your own email
   - Check if you receive a test version
   - Then use **Tasks → Preview & Send Tasks** for real send

### Optional: Add time-based trigger
- In Apps Script editor: **Triggers** (clock icon) → **Add Trigger**
- Choose `autoSendTasks` function
- Event source: Time-driven
- Type: Week timer → e.g. Every Monday at 7–8 AM

### Troubleshooting
- No menu? Run `onOpen()` manually
- Email quota error? Gmail allows ~100 emails/day for free accounts
- Errors? Check the **Logs** sheet or run functions step-by-step

For questions: pavlo.tsyhanash@gmail.com or Telegram @R_O_R_K
