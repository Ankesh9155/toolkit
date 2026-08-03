# Suppression & TAL Toolkit

A FastAPI-powered, browser-based version of the 5-tool suppression/TAL workflow
(Email Scrubber, TAL:SUPP Scrubber, TAL:SUPP Overlap Analysis, List Combine &
Split, List Comparison).

## Setup

```bash
cd toolkit
pip install -r requirements.txt --break-system-packages
uvicorn main:app --reload --port 8000
```

Open **http://127.0.0.1:8000** in your browser.

## Tools

1. **Email Scrubber** — Upload a suppression file (MD5-hashed and/or plain
   emails, mixed OK) and a lead list. Outputs clean vs. suppressed leads.
   Matching works both ways: a plain lead email is suppressed if it (or its
   MD5 hash) appears anywhere in the suppression file.

2. **TAL : SUPP Scrubber** — Upload an account/domain suppression file and a
   Target Account List. Domain/website/company columns are auto-detected.
   Outputs clean vs. suppressed accounts.

3. **TAL : SUPP Overlap Analysis** — Upload a TAL plus two suppression
   sources (account-level and email-level). Appends two boolean columns
   ("On Account Suppression", "On Email Suppression") to the TAL and
   downloads the enriched file.

4. **List Combine & Split**
   - *Combine*: upload any mix of files (CSV/XLSX with multiple tabs/TXT).
     Email-like columns are pooled and deduped into one email list; the rest
     is treated as account rows and deduped by company name + domain when
     available.
   - *Split*: upload one file, choose a column, and get a ZIP with one CSV
     per unique value in that column.

5. **List Comparison** — Compare 2 or 3 lists. Reports total unique values,
   how many appear in all lists / exactly 2 / exactly 1, and downloads a
   full breakdown spreadsheet.

## Notes

- All processing happens on your machine (localhost) — no data leaves your
  computer.
- Column auto-detection looks for column names containing "email",
  "domain", "website", etc., and falls back to sampling values against
  email/domain regex patterns.
- Generated files are held in memory only for download and are cleared when
  the server restarts.
