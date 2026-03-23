# ROUTR_SHOWCASE

Public-safe desktop automation platform for high-volume dispatch workflows.

> This repository is view-only and intended for technical review. It is not distributed as a runnable production build.

## In 10 Seconds

**What is this?**  
`ROUTR_SHOWCASE` is a Python desktop application that automates repetitive dispatch operations using a CustomTkinter UI + Selenium workflow engine.

**What problem does it solve?**  
It reduces manual effort in route/message handling workflows where teams otherwise click through repetitive UI tasks, copy data, and process exceptions by hand.

**What tech is used?**  
Python, CustomTkinter/Tkinter, Selenium WebDriver, requests, pandas, threaded background workers, and structured logging/retry patterns.

**What’s the impact?**  
It standardizes operational execution, improves consistency, and helps teams process high-volume dispatch activity faster with fewer manual errors.

## Key Features

- Automated message scanning and action routing
- Route extraction, validation, and workflow automation
- Auto-reply and auto-clear operational flows
- Real-time desktop dashboard with logs, counters, and controls
- API + browser automation hybrid processing
- Defensive fallbacks for dynamic web pages and transient UI failures

## Tech Stack

- Python
- Selenium WebDriver
- CustomTkinter / Tkinter
- requests
- pandas
- threading + structured logging

## Impact

- Reduces repetitive manual dispatch processing
- Improves operational consistency and response speed
- Decreases error-prone copy/paste and UI navigation work
- Makes high-volume workflow execution more scalable for operations teams

## Proof of Work

### Screenshots

Add project visuals here:

- Dispatch dashboard UI
- Automation run in progress
- Route/action output view

Create a `screenshots/` folder and embed images:

```markdown
![Dispatch Dashboard](screenshots/dashboard.png)
![Automation Run](screenshots/automation-run.png)
![Route Output](screenshots/route-output.png)
```

### Demo Video

Add a short walkthrough (2-5 minutes):

- Start app
- Show monitor/actions running
- Show logs/status and outcomes

Link format:

```markdown
[Watch Demo](https://your-demo-link)
```

### Case Study (Problem -> Solution -> Result)

**Problem**  
Dispatch workflows required repetitive manual processing, causing delays and inconsistency.

**Solution**  
Built a Python automation platform with a desktop control UI and Selenium-driven workflow handling for route/message operations.

**Result**  
Faster execution, better operational consistency, and less manual overhead for recurring dispatch tasks.

## Project Structure

```text
ROUTR_SHOWCASE/
├── ROUTR_SHOWCASE.py
├── requirements.txt
├── screenshots/
└── README.md
```

## Public Safety / Sanitization

To make this repository safe for public visibility:

- Removed hardcoded API keys and moved config to environment variables
- Removed private access allowlists and internal identifiers
- Replaced internal production URLs with placeholders
- Replaced company/provider branding terms with neutral placeholders
- Excluded credentials, local logs, generated artifacts, and private datasets via `.gitignore`

## Notes

- This repository is intended for technical review.
- Enterprise integrations are anonymized/stubbed where needed for safe public sharing.
