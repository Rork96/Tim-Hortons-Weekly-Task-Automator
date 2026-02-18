# Tim Hortons Weekly Task Automator

Google Apps Script tool that **automatically collects tasks** from multiple Google Sheets tabs and **sends personalized, beautifully formatted HTML emails** to restaurant managers every week.

**Real-world production tool** used at Tim Hortons restaurants in Winnipeg — reduces weekly manual preparation from 3–6 hours to ~2–3 minutes, eliminates distribution errors, and provides managers with clean, branded, calendar-style task overviews.

<p align="center">
  <img src="demo/sample-email-full.png" alt="Full example of personalized task email sent to a manager" width="600">
</p>

## Demo Spreadsheet (Template)

Test the automation yourself with this shared template:

<p align="center">
  <a href="https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing">
    <img src="https://img.shields.io/badge/Demo%20Spreadsheet-Open%20in%20Google%20Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white" alt="Open Demo Sheet">
  </a>
</p>

- Click above → File → Make a copy to your Drive
- Paste code from `/src/` into Extensions → Apps Script
- Run `onOpen()` to initialize
- Test with **Send Test Emails** first!

**Note:** This is a sanitized demo with placeholder data. Real emails/names are removed for privacy.

## Key Features

- Aggregates tasks from 10+ dedicated sheets:
  - Leaders Board (shift planning with visual calendar)
  - Self-Evaluations
  - Calibration & Receiving Logs
  - Weekly Reports - Section Assignments
  - CLUSTER - Training Videos
  - Supervisor Task Assignments
  - Deadlines Submissions
  - Settings (managers & emails)
  - Announcements (global messages)
  - Logs (full audit trail)
- Beautiful two-week Leaders Board table in emails (with ✓ checkmarks)
- Global announcements included in every email
- **Test mode** — send all personalized emails to your own address (Send Test Emails)
- **Preview mode** — shows summary of all tasks per manager before sending (Preview & Send Tasks)
- Skip any manager by setting email to "-" in Settings
- Gmail quota check + automatic warning emails to admin on issues
- Custom **Tasks** menu: preview, test, auto-send, instructions, announcements & logs
- Supports scheduled auto-run (e.g. every Monday via trigger)

## Technologies

- Google Apps Script (JavaScript)
- Google Sheets as structured database
- GmailApp for reliable email sending
- HTML + inline CSS for responsive, branded email templates
- Spreadsheet UI (custom menu, alerts, dialogs)

## Screenshots

### Spreadsheet Tabs Overview
All required input, config and log sheets.

![Spreadsheet Tabs List](demo/Spreadsheet-Tabs-List.png)  
*(Screenshot shows partial list; full set includes Settings, Leaders Board, Self-Evaluations, Calibration & Receiving Logs, Weekly Reports - Section Assignments, CLUSTER - Training Videos, Deadlines Submissions, Supervisor Task Assignments, Announcements, Logs)*

### Custom Tasks Menu
User-friendly menu with all available options.

![Tasks Menu](demo/tasks-menu.png)

### Managers & Emails Settings
Configure recipients (use "-" to skip).

![Settings Sheet](demo/Settings.png)

### Announcements Sheet
Add messages visible to all managers.

![Announcements Sheet](demo/announcements-sheet.png)

### Leaders Board Input Example
Source data parsed into calendar table in emails.

![Leaders Board Example](demo/leaders-board-example.png)

### Sample Received Email (Full View)
What managers receive — branded, structured, easy to read.

<p align="center">
  <img src="demo/sample-email-full.png" alt="Personalized weekly tasks email with Tim Hortons branding, Leaders Board calendar, task lists and announcements" width="600">
</p>

### Additional Input Sheets Examples
- Self-Evaluations: ![Self-Evaluations](demo/Self-Evaluations.png)
- Calibration & Receiving Logs: ![Calibration & Receiving Logs](demo/Calibration%20&%20Receiving%20Logs.png)
- Supervisor Task Assignments: ![Supervisor Task Assignments](demo/Supervisor%20Task%20Assignments.png)
- Deadlines Submissions: ![Deadlines Submissions](demo/Deadlines%20Submissions.png)
- Logs (audit trail): ![Logs Sheet](demo/Logs.png)

## Installation & Usage

1. Make a copy of the demo spreadsheet (or use your own)
2. Go to **Extensions → Apps Script**
3. Delete default code and paste files from `/src/` (or combine into one file)
4. Save the project
5. Reload the spreadsheet — **Tasks** menu appears
6. **Test first**: Tasks → Send Test Emails → enter your email
7. **Send for real**: Tasks → Preview & Send Tasks → review → confirm

Detailed guide → [INSTALL.md](INSTALL.md)

## Business Impact

- Weekly preparation time: reduced from ~4 hours → under 5 minutes
- Zero manual copy-paste or assignment errors
- Managers get clean, branded HTML emails (calendar + lists) instead of plain text/screenshots
- Full audit trail in Logs sheet for accountability
- Tested in real restaurant operations (Winnipeg, Manitoba)

## License

MIT License — see [LICENSE](LICENSE) for full text.

Made with ❤️ in Winnipeg, Manitoba  
Author: Pavlo Tsyhanash  
Email: pavlo.tsyhanash@gmail.com  
Telegram: https://t.me/R_O_R_K

Open to collaborations — happy to adapt for other restaurants, retail or operations teams!
