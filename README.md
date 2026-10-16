# College Outreach Campaign Tracker

A spreadsheet I built to run a college outreach campaign: reaching Training & Placement Officers (TPOs) and students, following up on time, and reporting how many students registered and applied.

## What's inside

| Sheet | What it's for |
|---|---|
| **Dashboard** | Key numbers, college funnel, response rate by channel (call / email / WhatsApp), city-wise results, follow-ups due |
| **Colleges** | One row per college: TPO, preferred channel, status, students reached, registrations, applications, auto follow-up date |
| **Outreach Log** | One row per call, email or WhatsApp message, with the outcome |
| **Daily Metrics** | Touches per channel, positive responses and new registrations each day, with a trend chart |
| **Templates** | Ready-to-use TPO email, call script, follow-up, student WhatsApp message and reminder |

## How it works

- Every total and rate is a formula, so logging a new call updates the dashboard straight away.
- A college marked *Contacted* or *Interested* gets a follow-up date 3 days after the last touch. If that date has passed, it's flagged in red.
- Channel response rates show which channel works best, so effort can shift toward it.

## Sample results (fictional data, 1–14 Oct 2026)

- 26 of 30 colleges contacted, 8 onboarded
- 103 outreach touches, 51% positive response rate
- WhatsApp had the best response rate (59%), email the lowest (38%)
- 4,308 students reached → 908 registrations (21%) → 370 applications

All colleges, contacts and numbers are made up for demonstration. Replace them with real data to use it.

## Rebuild

```
pip install openpyxl
python build_tracker.py
```
