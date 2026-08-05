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
- Generated files are written to a temp directory on disk (not kept as bytes
  in memory) and are cleared after 20 minutes or when the server restarts.
- Uploads over 9MB (.xlsx/.xlsm/.csv) are parsed in row batches (20,000 rows
  at a time) instead of being loaded whole, so peak memory during *parsing*
  stays bounded on larger files. Output is identical either way — this only
  changes how the file is read in, not what comes out.
- Uploads are capped at 200MB by default (`MAX_UPLOAD_MB` env var to
  change), rejected with a clean 413 instead of exhausting memory mid-parse.
  This cap does not by itself make large files "safe" — see "Memory sizing"
  below.

## Memory sizing for large files

Row-batch parsing only bounds the *parsing* step; once a file is parsed it
still becomes one full pandas DataFrame in memory, typically **3–8x the
raw file size** (worse for .xlsx than .csv, due to per-cell Python object
overhead). On top of that, each tool builds one or more full-size Excel
outputs before they're written to disk. So for a single request, peak
memory is roughly:

    raw upload bytes + (3-8x that) as a DataFrame + one Excel output buffer

A 40MB upload can peak north of 300-400MB; a 200MB upload can peak well
over 1GB. Render's Starter instance (512MB RAM) cannot reliably process
uploads much above ~40-60MB no matter how the code chunks the read —
**the fix for the top of a 1-200MB range is more RAM, not more chunking.**
If you need to reliably support uploads up to 200MB, move the Render
service to a plan with at least 2GB RAM (Standard tier or higher).
