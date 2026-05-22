# Okta Profile Update Metrics

A Python script that queries the Okta System Log API to pull `user.account.update_profile` 
events and calculate estimated Event Hook trigger volume.

## What It Does
- Pulls all profile update events from Okta PROD for a defined date range
- Calculates total events, daily average, and peak day
- Exports results to a formatted Excel file with three tabs:
  - **Summary** — key metrics at a glance
  - **Daily Breakdown** — event count per day
  - **Raw Events** — full event log with actor, target user, IP, and outcome

## Setup

1. Clone the repo
2. Create a virtual environment and install dependencies:
```bash
   pip install requests openpyxl python-dotenv
```
3. Create a `.env` file in the project root: