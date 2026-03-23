# ROUTR_SHOWCASE

Public-safe showcase edition of a route automation desktop tool for portfolio and resume use.

## What This Version Is

This repository is a sanitized copy prepared for public visibility. It demonstrates:

- Desktop automation patterns with Selenium
- Operational workflow UI built with CustomTkinter
- Route and dispatch-oriented data handling
- Error handling, retries, and long-running automation management

## Public Safety Changes

The showcase copy removes or neutralizes private/internal data:

- Removed hardcoded API keys
- Removed private access allowlists
- Replaced internal production URLs with placeholder endpoints
- Added strict `.gitignore` patterns for secrets and local artifacts
- Excluded company-specific credentials, logs, and local datasets

## Setup

1. Create and activate a virtual environment
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Configure environment variables (example):

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

- This repository is intended for demonstration and code review.
- End-to-end behavior that depends on private enterprise systems is intentionally stubbed/anonymized.
- All branding and internal identifiers in code/UI text were replaced with neutral placeholders.
