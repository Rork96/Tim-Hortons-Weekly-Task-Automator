# Installation & Setup Guide

This Google Apps Script automates weekly task distribution for Tim Hortons managers.

### Prerequisites
- Google account
- Google Sheet with required tabs — start with the [demo template](https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing)
- Basic Google Sheets knowledge

### Step-by-Step Setup

1. **Prepare the Sheet**
   - Open [demo](https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing)
   - File → Make a copy to your Drive
   - Verify tabs: Settings, Announcements, Leaders Board, Self-Evaluations, Calibration & Receiving Logs, Weekly Reports - Section Assignments, CLUSTER - Training Videos, Supervisor Task Assignments, Deadlines Submissions

2. **Open Apps Script**
   - In the Sheet: **Extensions → Apps Script**

3. **Add the Code (Single-File Approach)**
   - Delete default `Code.gs`
   - Paste the entire script code (from `/src/main.gs` or combined file in this repo)
   - **Note:** Everything is in one file for simplicity and fast deployment. Modular structure is optional.

4. **Save & Authorize**
   - Save project
   - Run `onOpen()` → authorize permissions (read/write Sheets + send emails)
   - Approve all scopes

5. **Initialize**
   - Reload the sheet tab
   - **Tasks** menu appears in top bar

6. **Test Safely**
   - **Tasks → Send Test Emails** → enter your email → verify receipt/format
   - **Tasks → Preview & Send Tasks** → check summary → confirm send

### Optional: Weekly Auto-Run
- Apps Script → Triggers → Add Trigger
- Function: `autoSendTasks`
- Event: Time-driven → Week timer → Monday 7–8 AM

### Troubleshooting
- No menu? Manually run `onOpen()`
- Quota error? Gmail free limit ~100/day — wait 24h
- Errors? Check **Logs** sheet or console logs
- Contact: pavlo.tsyhanash@gmail.com | Telegram @R_O_R_K

### Customization & Contribution
- Add new task sources: extend `collectTasks()` function
- Change email styling: edit HTML template in `sendEmails()`
- Multi-language? Add toggle for Ukrainian/English
- Pull requests welcome for improvements!

Happy automating!
