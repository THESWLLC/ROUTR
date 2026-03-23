# ROUTR_SHOWCASE

Public-safe desktop automation platform for high-volume dispatch workflows.

> This repository is a sanitized technical showcase for code review and portfolio evaluation. It is intentionally non-runnable and excludes private integrations, credentials, internal endpoints, and production datasets.

## In 10 Seconds

**What is this?**  
`ROUTR_SHOWCASE` is a Python desktop application that automates repetitive dispatch operations using a CustomTkinter UI + Selenium workflow engine.

**What problem does it solve?**  
It reduces manual effort in route/message handling workflows where teams otherwise click through repetitive UI tasks, copy data, and process exceptions by hand.

**What tech is used?**  
Python, CustomTkinter/Tkinter, Selenium WebDriver, requests, pandas, threaded background workers, and structured logging/retry patterns.

**What’s the impact?**  
It standardizes operational execution, improves consistency, and helps teams process high-volume dispatch activity faster with fewer manual errors.

## What This Demonstrates

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

## Architecture Overview

- **UI Layer:** Desktop control panel and monitoring views using CustomTkinter/Tkinter
- **Automation Layer:** Selenium-based browser workflow engine for operational tasks
- **Data/Integration Layer:** API calls, parsing, validation, and action routing logic
- **Reliability Layer:** Retry loops, timeout guards, stale element recovery, and structured logging

## Workflow Overview

1. Authenticate into an operational session (sanitized in public version)
2. Monitor high-volume message/task queues
3. Extract route/context data and classify actionable items
4. Trigger automated actions (reply, clear, route workflows)
5. Emit status logs/counters for operator visibility and auditability

## Impact

- Reduced repetitive manual touchpoints in recurring dispatch workflows (estimated 30-40% improvement)
- Improved operational consistency through standardized rule-based actions and response flows
- Decreased error-prone copy/paste and repeated UI navigation in high-volume message handling
- Increased throughput for repetitive route/message tasks during peak operational periods

## Proof of Work

### Code Evidence

- `ROUTR_SHOWCASE.py` contains end-to-end UI, workflow, and automation orchestration.
- The structure demonstrates production-style defensive coding for real-world dynamic web flows.

### Visual Evidence

Screenshots and walkthrough clips are maintained separately to keep this repo focused on sanitized technical artifacts.

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
