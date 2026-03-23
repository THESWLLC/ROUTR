# ROUTR_SHOWCASE

Public-safe desktop automation platform for high-volume dispatch workflows.

## In 10 Seconds

**What is this?**  
`ROUTR_SHOWCASE` is a Python desktop application that automates repetitive dispatch operations using a CustomTkinter UI + Selenium workflow engine.

**What problem does it solve?**  
It reduces manual effort in route/message handling workflows where teams otherwise click through repetitive UI tasks, copy data, and process exceptions by hand.

**What tech is used?**  
Python, CustomTkinter/Tkinter, Selenium WebDriver, requests, pandas, threaded background workers, and structured logging/retry patterns.

**What’s the impact?**  
It standardizes operational execution, improves consistency, and helps teams process high-volume dispatch activity faster with fewer manual errors.

## Why This Project Matters

`ROUTR_SHOWCASE` demonstrates production-style automation engineering that combines:

- reliable browser automation for operational workflows
- a desktop UI that non-technical users can run daily
- defensive engineering (retry logic, stale element recovery, timeout handling)
- data normalization and workflow orchestration across multiple tasks

This is a sanitized public version of a real-world operations tool.

## Core Capabilities

- **Automated Workflow Monitoring:** Continuously monitors dispatch-style message queues and reacts to actionable items.
- **Smart Route Handling:** Extracts route data from messages, validates content, and triggers automated follow-up actions.
- **Auto-Reply / Auto-Clear Flows:** Implements response and cleanup workflows with safeguards to reduce manual repetitive work.
- **Operational Dashboard UI:** Multi-panel CustomTkinter interface with status indicators, counters, logs, and controls.
- **API + UI Hybrid Processing:** Uses API calls for package/exception data and Selenium-driven UI automation where required.
- **Resilience Under Change:** Includes robust fallbacks for dynamic pages (stale references, click interception, delayed rendering).

## Technical Highlights

- **Language:** Python
- **UI:** CustomTkinter / Tkinter
- **Automation:** Selenium WebDriver
- **Data Handling:** pandas, regex parsing, request/response processing
- **Concurrency:** Threaded workers + UI-safe callbacks
- **Reliability Patterns:** retries, timeout guards, structured logging, exception recovery loops

## Architecture Snapshot

- `ROUTR_SHOWCASE.py` contains:
  - authentication/session navigation helpers
  - message scanning and pattern-based action routing
  - route and dispatch workflow handlers
  - operational dashboard rendering and user controls
  - background worker loops and status/reporting hooks

## Resume-Relevant Engineering Skills Demonstrated

- Building end-user automation products, not just scripts
- Translating operational pain points into deterministic software workflows
- Balancing speed and safety in long-running browser automation
- Designing UI-first tooling for team adoption
- Refactoring private/internal systems into compliant public showcases

## Public Safety / Sanitization

To make this repository safe for public visibility:

- Removed hardcoded API keys and moved config to environment variables
- Removed private access allowlists and internal identifiers
- Replaced internal production URLs with placeholders
- Replaced company/provider branding terms with neutral placeholders
- Excluded credentials, local logs, generated artifacts, and private datasets via `.gitignore`

## Quick Start

1. Create and activate a virtual environment
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Configure environment variables:

```bash
set GOOGLE_API_KEY=your_key_here
set ROUTR_API_BASE=https://api.example.com
set ROUTR_UI_BASE=https://app.example.com
```

4. Run:

```bash
python ROUTR_SHOWCASE.py
```

## Notes

- This repository is intended for technical review and discussion.
- Enterprise integrations are anonymized/stubbed where needed for safe public sharing.
