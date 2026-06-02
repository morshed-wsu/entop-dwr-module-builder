# enTop Clients DWR Engine — Web App Builder

> **Version:** 1.5.0 Reporting Framework  
> **Company:** enTop (pi Cyber Logic Ltd.)  
> **Live App:** https://entopbd.com/apps/  
> **Platform:** Namecheap Stellar Shared Hosting · cPanel Python App (v3.11.15)

---

## Overview

The **enTop Clients DWR Engine** is a lightweight Flask-based web application that automates the generation of **SDL Daily Work Reports (DWR)** directly on the company's own shared hosting server — ensuring full data privacy for client production data.

The system accepts a daily Excel input file (`DD-MM-YYYY_SDL_Daily_Work_Report.xlsx`), processes it through a four-step report generation pipeline, and produces professionally formatted `.docx` output files that match enTop's branded Word templates.

The app is deployed at `https://entopbd.com/apps/` and operated via a browser-based form interface.

---

## Project Background

Report generation was previously handled by AI platforms (Claude/Perplexity) using `CLAUDE.md` and `RUNBOOK.md` instruction files. This project migrates equivalent logic into Python scripts running server-side, eliminating the need to share raw client data with external AI services.

The `CLAUDE.md` master instructions were split into four step-specific `.md` files to keep each processor lean and compatible with shared hosting memory constraints.

---

## App Structure

```
entopbd.com/apps/
├── entop_sdl_app.py                        # Main Flask app (entry point)
├── passenger_wsgi.py                       # Passenger WSGI entry
├── daily_internal_qc_entop_sdl_prod.py    # QC jobs script
├── draft_sdl_daily_report.py              # Legacy daily report drafter
├── ecedb_data_entry.py                    # Data entry helper
├── sdl_dwr_step1_generator.py             # Step 1 DWR generator
├── sdl_dwr_step2_generator.py             # Step 2 DWR generator
├── sdl_dwr_step3_generator.py             # Step 3 DWR generator
├── sdl_dwr_step4_generator.py             # Step 4 DWR generator
├── Daily Work Report Step 1 Template - enTop v1.0.0.docx
├── Daily Work Report Step 2 Template - enTop v1.0.0.docx
├── Daily Work Report Step 3 Template - enTop v1.0.0.docx
├── Daily Work Report Step 4 Template - enTop v1.0.0.docx
├── DWR_Cover_Page.docx                    # Cover page template
├── enTop-SDL_team_members.xlsx            # Team member reference data
├── requirements.txt
├── uploads/                               # Uploaded Excel files
│   └── dwr_bind_steps/                    # Step files for binding
├── generated_reports/                     # Output .docx files
└── public/                                # Static assets
```

---

## Web App Pages

### Home — `https://entopbd.com/apps/`
The **enTop Clients – Service Delivery Hub** home page provides the following actions:
- **SDL Daily QC Jobs** — upload an Excel file and run internal QC
- **Draft SDL Daily Report** — upload an Excel file to generate a draft report
- **Run Script 3** — utility script
- **Go to Daily Work Report** — navigates to the 4-step DWR generation form

### Daily Work Report — `https://entopbd.com/apps/daily-work-report`
The SDL DWR generation page where operators:
1. Upload the daily Excel file (`DD-MM-YYYY_SDL_Daily_Work_Report.xlsx`)
2. Generate each step independently using the colour-coded step buttons
3. Bind all four generated Word files into a single final DWR document

---

## DWR Generation Pipeline

Each step reads the **same uploaded Excel file** and produces a standalone `.docx` output. Steps must be run in order before binding.

| Step | Content | Generator Script | Template |
|------|---------|-----------------|---------|
| Step 1 | Operational Tables (completed/incomplete jobs, resources) | `sdl_dwr_step1_generator.py` | `Daily Work Report Step 1 Template - enTop v1.0.0.docx` |
| Step 2 | Daily Metrics at a Glance, Top Ordering & Quoted Customers | `sdl_dwr_step2_generator.py` | `Daily Work Report Step 2 Template - enTop v1.0.0.docx` |
| Step 3 | Leaderboard, Descriptive Statistics, Histogram Chart | `sdl_dwr_step3_generator.py` | `Daily Work Report Step 3 Template - enTop v1.0.0.docx` |
| Step 4 | Key Observations | `sdl_dwr_step4_generator.py` | `Daily Work Report Step 4 Template - enTop v1.0.0.docx` |

### Output Filename Format
```
DWR_YYYY-DWR#_DD-Mmm-YYYY-Step_N.docx
```
Example: `DWR_2026-91_02-Jun-2026-Step_1.docx`

---

## Input File

**File naming convention:**
```
DD-MM-YYYY_SDL_Daily_Work_Report.xlsx
```

**Required sheets:**
- `Summary` — DWR number (B1), report date (C1), resource utilisation rows (rows 3–6), summary aggregates
- `Production_Center` — Per-job records starting at row 8 with columns for job type, job number, value, customer, group work flag, remaining time, and status

**Key status codes used:**
- `C` — Completed
- `IP` — In Progress
- `NA` — Not Attempted
- `GU` — Given Up

---

## Key Business Logic

### Revenue Frequency Distribution
Orders and quotes are bucketed into value brackets:
`$0–500 | $501–2,000 | $2,001–5,000 | $5,001–10,000 | $10,001–15,000 | $15,001–25,000 | $25,001–50,000 | $50,001+`

### Top Ordering / Quoted Customer Analysis
Customers are ranked by order value, quote value, job count, and remaining time. The top 3 customers are shown individually, with all remaining customers aggregated as **"OTHER CUSTOMERS"**.

### Resource Capacity Tracking (40-Man-Hour)
Resource utilisation is tracked against a 40-man-hour capacity benchmark using data from the `Summary` sheet rows 3–6, columns C–I.

---

## Bind DWR Step Files

After generating all four step `.docx` files, they can be merged into a single Final DWR document via the **Bind DWR Step Files** section on the daily-work-report page. Files are uploaded in Step 1 → Step 2 → Step 3 → Step 4 order and concatenated into one report.

---

## Instruction Reference Files (Space Sources)

The generation logic for each step is governed by the following instruction files stored in the Perplexity Space:

| File | Purpose |
|------|---------|
| `sdl_dwr_step1_generator.md` | Step 1 Python script (operational tables) |
| `sdl_dwr_step2_generator.md` | Step 2 Python script (metrics & customer analysis) |
| `sdl_dwr_step3_generator.md` | Step 3 Python script (leaderboard & histogram) |
| `sdl_dwr_step4_generator.md` | Step 4 Python script (key observations) |
| `entop_sdl_app.md` | Main Flask app source |
| `requirements.txt` | Python dependencies |

---

## Requirements

```
openpyxl
python-docx
lxml
flask
pandas
numpy
werkzeug
pymysql
mysql-connector-python
docxcompose
```

---

## Deployment Notes

- **Hosting:** Namecheap Stellar Shared Hosting
- **Python Version:** 3.11.15 (via cPanel Setup Python App)
- **App Root:** `/home/encabitr/entopbd.com/apps`
- **App URI:** `apps.entopbd.com/`
- **WSGI Entry:** `passenger_wsgi.py`
- All template `.docx` files must be present in the app root directory before running any step generator.
- The `uploads/` and `generated_reports/` directories are created automatically on first run.
- After any code update, restart the Python app from cPanel → Python Apps → Restart (↺ icon).

---

## Workflow Summary

```
Upload Excel File
      │
      ▼
Generate Step 1 DWR  →  DWR_YYYY-#_DD-Mmm-YYYY-Step_1.docx
Generate Step 2 DWR  →  DWR_YYYY-#_DD-Mmm-YYYY-Step_2.docx
Generate Step 3 DWR  →  DWR_YYYY-#_DD-Mmm-YYYY-Step_3.docx
Generate Step 4 DWR  →  DWR_YYYY-#_DD-Mmm-YYYY-Step_4.docx
      │
      ▼
Bind All 4 Steps  →  Final DWR Document (.docx)
```

---

## Maintained By

**enTop · pi Cyber Logic Ltd.**  
Perplexity Space: *enTop Clients DWR Engine Web App Builder*
