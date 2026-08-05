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
- Uploads over 9MB (.xlsx) use openpyxl's read-only row iterator instead of
  the default engine, which parses the whole workbook's XML into memory up
  front. Output is identical either way — this only changes how the file is
  read in, not what comes out.
- Uploads are capped at 40MB by default (`MAX_UPLOAD_MB` env var to change),
  rejected with a clean 413 instead of exhausting memory mid-parse — sized
  for Render Free's 512MB RAM. See "Memory sizing" below.
- At most one upload is parsed/transformed at a time (`MAX_CONCURRENT_JOBS`
  env var, default 1). Extra requests queue instead of running in parallel,
  since concurrent large uploads is the fastest way to multiply peak memory
  past what one request alone would use.

## Memory sizing for large files

Every tool loads its input into a pandas DataFrame, which typically runs
**3–8x the raw file size** in memory (worse for .xlsx than .csv, due to
per-cell Python object overhead), then builds one or more full-size Excel
outputs from it. So for a single request, peak memory is roughly:

    raw upload bytes + (3-8x that) as a DataFrame + one Excel output buffer

A handful of optimizations keep that peak as low as it can go without
changing what any tool outputs:
- Chunked reads build the file's rows once instead of building N partial
  DataFrames and concatenating them (concat briefly holds both copies at
  once — see comments in `logic.py`'s `_read_xlsx_chunked`/`_read_csv_chunked`).
- MD5 hashes in the suppression set are stored as raw bytes instead of hex
  strings (~35-40% less memory for that structure).
- Requests are serialized (`MAX_CONCURRENT_JOBS=1`) so peak memory reflects
  one upload at a time, not however many arrive together.
- `gc.collect()` runs at specific points where openpyxl/xlsxwriter objects
  with reference cycles (worksheet ↔ parent workbook, etc.) go out of scope
  — those need the cyclic collector, not just refcounting, to actually free.

None of this changes the underlying math: a 40MB upload can still peak
north of 150–300MB depending on file shape, which is why the default cap
is 40MB on a 512MB instance, not higher. Going meaningfully past that
without more RAM would require processing rows in a streaming pass instead
of building a DataFrame at all — a real rewrite of the processing pipeline,
not a tuning change. If you need to reliably support uploads much larger
than 40MB, either that rewrite or moving the Render service to a plan with
more RAM are the two real options.
