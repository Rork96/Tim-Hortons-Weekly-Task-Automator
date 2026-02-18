![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google&logoColor=white)
![MIT License](https://img.shields.io/badge/License-MIT-green)
![Real World Usage](https://img.shields.io/badge/Used%20at-Tim%20Hortons-orange)

# Tim Hortons Weekly Task Automator

Google Apps Script tool that **automatically collects tasks** from multiple Google Sheets tabs and **sends personalized, beautifully formatted HTML emails** to restaurant managers every week.

**Production-grade automation** deployed at Tim Hortons restaurants in Winnipeg — reduces weekly manual ops time from 3–6 hours to under 5 minutes, eliminates human errors, and delivers clean, branded emails with visual calendars.

<p align="center">
  <img src="demo/sample-email-full.png" alt="Full example of personalized task email sent to a manager" width="600">
</p>

### Problem → Solution
Manual weekly task assignment across 10+ sheets for 10+ managers took hours of copy-paste, prone to mistakes and oversights.  
This tool fully automates data aggregation, formatting, preview/testing, and sending — with zero-touch options via triggers.

## Demo Spreadsheet (Template)

Quickly test it yourself:

<p align="center">
  <a href="https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing">
    <img src="https://img.shields.io/badge/Demo%20Spreadsheet-Open%20in%20Google%20Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white" alt="Open Demo Sheet">
  </a>
</p>

- Click above → **File → Make a copy** to your Drive
- Paste the script (single file) into **Extensions → Apps Script**
- Run `onOpen()` to initialize
- Test safely with **Send Test Emails**

**Note:** Sanitized demo with placeholder data. Real emails/names removed for privacy.

## Key Features

- Aggregates tasks from 10+ sheets (Leaders Board, Self-Evaluations, Calibration Logs, Weekly Reports, CLUSTER Training, Supervisor Tasks, Deadlines, Settings, Announcements, Logs)
- Two-week Leaders Board rendered as clean calendar table with ✓ checkmarks
- Global announcements in every email
- **Safe testing**: Send Test Emails → all to your inbox
- **Preview before send**: Summary dialog per manager
- Skip managers via "-" in Settings
- Gmail quota monitoring + admin alerts
- Custom **Tasks** menu (preview, test, auto, instructions, logs)
- Trigger-ready for weekly auto-send (Monday mornings)

## Technologies

- Google Apps Script (JavaScript, single-file implementation)
- Google Sheets as structured data source
- GmailApp for sending
- HTML + inline CSS for branded, responsive emails
- Spreadsheet UI (custom menu, alerts, dialogs)

## Screenshots

### Spreadsheet Tabs Overview
Full structure of input, config and log sheets.

![Spreadsheet Tabs List](demo/Spreadsheet-Tabs-List.png)

### Custom Tasks Menu
All available actions at a glance.

![Tasks Menu](demo/tasks-menu.png)

### Managers & Emails Settings
Recipient configuration.

![Settings Sheet](demo/Settings.png)

### Announcements Sheet
Global broadcast messages.

![Announcements Sheet](demo/announcements-sheet.png)

### Leaders Board Input Example
Source for calendar rendering.

![Leaders Board Example](demo/leaders-board-example.png)

### Sample Received Email (Full View)
What managers actually get.

<p align="center">
  <img src="demo/sample-email-full.png" alt="Personalized weekly tasks email with Tim Hortons branding, Leaders Board calendar, task lists and announcements" width="600">
</p>

### Additional Input Sheets Examples (click to expand)
<details>
<summary>Click to see more sheet examples</summary>

- Self-Evaluations: ![Self-Evaluations](demo/Self-Evaluations.png)
- Calibration & Receiving Logs: ![Calibration & Receiving Logs](demo/Calibration%20&%20Receiving%20Logs.png)
- Supervisor Task Assignments: ![Supervisor Task Assignments](demo/Supervisor%20Task%20Assignments.png)
- Deadlines Submissions: ![Deadlines Submissions](demo/Deadlines%20Submissions.png)
- Logs (audit trail): ![Logs Sheet](demo/Logs.png)

</details>

## Installation & Usage

1. Make a copy of the [demo spreadsheet](https://docs.google.com/spreadsheets/d/1hdgul5uaL5jyhZzR-JLbXlIVdrX2p3MnRrzVagGT4gQ/edit?usp=sharing)
2. **Extensions → Apps Script**
3. Delete default code → paste the full script (single file from `/src/`)
4. Save → Run `onOpen()` → authorize (Sheets + Gmail)
5. Reload sheet → **Tasks** menu appears
6. **Test first**: Tasks → Send Test Emails → your email
7. **Send real**: Tasks → Preview & Send Tasks → review → confirm

Full guide → [INSTALL.md](INSTALL.md)

## Business Impact

- Weekly ops time: ~4 hours → <5 minutes
- 100% accuracy in task assignment
- Branded, readable emails replace messy text/screenshots
- Complete audit trail in Logs
- Proven in live restaurant environment (Winnipeg, Manitoba)

## License

MIT License — see [LICENSE](LICENSE)

Made with ❤️ in Winnipeg, Manitoba  
Author: Pavlo Tsyhanash  
Email: pavlo.tsyhanash@gmail.com  
Telegram: https://t.me/R_O_R_K

Open to forks, collaborations, or adaptations for other chains/operations!
