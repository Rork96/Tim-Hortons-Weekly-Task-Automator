# Installation & Setup Guide

This Google Apps Script automates weekly task distribution for Tim Hortons managers using Google Sheets.

### Prerequisites
- Google account
- Google Sheet with required tabs (Settings, Announcements, Logs, Leaders Board, etc.) — use the [demo template](https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing)
- Basic familiarity with Google Sheets

### Step-by-Step Setup

1. **Prepare your Google Sheet**
   - Open the [demo spreadsheet](https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing) or your own
   - File → Make a copy (to your Drive)
   - Ensure tabs exist: Settings, Announcements, Leaders Board, Self-Evaluations, Calibration & Receiving Logs, Weekly Reports - Section Assignments, CLUSTER - Training Videos, Supervisor Task Assignments, Deadlines Submissions

2. **Open Apps Script editor**
   - In the Sheet: **Extensions → Apps Script**

3. **Add the code**
   - Delete default `Code.gs`
   - Create files via + button (recommended structure):
     - `main.gs` — menu & main functions
     - `setup.gs` — sheet setup functions
     - `email-builder.gs` — HTML email generation
     - `data-collectors.gs` — task collection logic
     - `utils.gs` — logging & helpers
   - Or paste everything into one file if simpler
   - Copy code from `/src/` in this repository

4. **Save and authorize**
   - Save project (disk icon)
   - Run `onOpen()` → authorize permissions (Sheets + Gmail)
   - Approve all requests

5. **Reload the Sheet**
   - Close and reopen browser tab
   - **Tasks** menu appears in top bar

6. **Test the tool**
   - **Tasks → Send Test Emails** → enter your email → check inbox
   - **Tasks → Preview & Send Tasks** → review summary → confirm send

### Optional: Auto-Run Every Week
- Apps Script editor → Triggers (clock icon) → Add Trigger
- Function: `autoSendTasks`
- Event: Time-driven → Week timer → Monday 7–8 AM

### Troubleshooting
- No menu? Run `onOpen()` manually
- Email quota exceeded? Free Gmail: ~100/day — wait till next day
- Errors? Check **Logs** sheet or console
- Still issues? Contact: pavlo.tsyhanash@gmail.com or Telegram @R_O_R_K

Happy automating!
