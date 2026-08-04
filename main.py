"""
Suppression & TAL Toolkit — FastAPI backend
=============================================

Run with:
    pip install -r requirements.txt --break-system-packages
    uvicorn main:app --reload --port 8000

Then open http://127.0.0.1:8000 in your browser.

5 tools, mirroring the reference site:
  1. Email Scrubber           - suppress a lead list against an MD5/plain email suppression file
  2. TAL : SUPP Scrubber      - suppress a TAL (accounts) against a domain suppression file
  3. TAL : SUPP Overlap       - enrich a TAL with two suppression sources (account + email domains)
  4. List Combine & Split     - merge multiple files with dedupe, or split one file by column
  5. List Comparison          - compare 2-3 lists (venn-style overlap)

All processing happens in-memory; nothing is written to disk except transient
result blobs kept in a process-local cache for download.
"""

import io
import time
import uuid
from typing import List, Optional

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import logic

app = FastAPI(title="Suppression & TAL Toolkit")

# In-memory store for generated result files, keyed by a random token.
# {token: {"filename": str, "bytes": bytes, "media_type": str, "created": float}}
# Entries are purged after _RESULT_TTL_SECONDS to avoid unbounded memory growth.
_RESULTS = {}
_RESULT_TTL_SECONDS = 20 * 60


def _purge_expired_results():
    now = time.time()
    expired = [t for t, item in _RESULTS.items() if now - item["created"] > _RESULT_TTL_SECONDS]
    for t in expired:
        del _RESULTS[t]


def _store(filename: str, data: bytes, media_type: str) -> str:
    _purge_expired_results()
    token = uuid.uuid4().hex
    _RESULTS[token] = {"filename": filename, "bytes": data, "media_type": media_type, "created": time.time()}
    return token


@app.get("/api/download/{token}")
def download(token: str):
    item = _RESULTS.get(token)
    if not item:
        raise HTTPException(404, "Result not found or expired. Re-run the tool.")
    return StreamingResponse(
        io.BytesIO(item["bytes"]),
        media_type=item["media_type"],
        headers={"Content-Disposition": f'attachment; filename="{item["filename"]}"'},
    )


async def _read(upload: UploadFile):
    content = await upload.read()
    return content, upload.filename


async def _resolve_input(upload: Optional[UploadFile], pasted_text: Optional[str], label: str):
    """Accept either an uploaded file OR a pasted block of text (one email/hash
    per line). Exactly one must be provided. Pasted text is treated as a .txt
    file internally so it goes through the same parsing path as an upload."""
    if upload is not None and getattr(upload, "filename", None):
        return await _read(upload)
    if pasted_text and pasted_text.strip():
        return pasted_text.encode("utf-8"), f"{label}_pasted.txt"
    raise HTTPException(400, f"Provide {label} as either a file upload or pasted text.")


# ---------------------------------------------------------------------------
# Tool 1: Email Scrubber
# ---------------------------------------------------------------------------

@app.post("/api/lead-columns")
async def lead_columns_endpoint(
    file: Optional[UploadFile] = File(None),
    text: Optional[str] = Form(None),
):
    content, name = await _resolve_input(file, text, "leads_file")
    try:
        df = logic.read_single_table(content, name)
    except ValueError as e:
        raise HTTPException(400, str(e))
    detected = logic.detect_email_column(df)
    return JSONResponse({"columns": list(df.columns), "detected": detected})


@app.post("/api/email-scrubber")
async def email_scrubber_endpoint(
    suppression_file: Optional[UploadFile] = File(None),
    leads_file: Optional[UploadFile] = File(None),
    suppression_text: Optional[str] = Form(None),
    leads_text: Optional[str] = Form(None),
    email_column: Optional[str] = Form(None),
    out_filename: str = Form("scrubbed_leads"),
):
    supp_content, supp_name = await _resolve_input(suppression_file, suppression_text, "suppression_file")
    lead_content, lead_name = await _resolve_input(leads_file, leads_text, "leads_file")
    try:
        result = logic.email_scrubber(supp_content, supp_name, lead_content, lead_name, email_column)
    except ValueError as e:
        raise HTTPException(400, str(e))

    clean_bytes = logic.df_to_excel_bytes({"Clean Leads": result["clean_df"]})
    supp_bytes = logic.df_to_excel_bytes({"Suppressed": result["suppressed_df"]})
    dup_bytes = logic.df_to_excel_bytes({"Duplicates": result["duplicates_df"]})
    all_bytes = logic.df_to_excel_bytes({"All Leads": result["annotated_df"]})
    clean_token = _store(f"{out_filename}_clean.xlsx", clean_bytes,
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    supp_token = _store(f"{out_filename}_suppressed.xlsx", supp_bytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    dup_token = _store(f"{out_filename}_duplicates.xlsx", dup_bytes,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    all_token = _store(f"{out_filename}_all_leads.xlsx", all_bytes,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    total = result["total_leads"]
    preview_limit = 500
    suppressed_only_df = result["annotated_df"][result["annotated_df"]["Found in suppression"] == "Yes"]
    return JSONResponse({
        "total_leads": total,
        "suppressed": result["suppressed"],
        "clean": result["clean"],
        "duplicates": result["duplicates"],
        "email_column_used": result["email_column_used"],
        "clean_download": f"/api/download/{clean_token}",
        "suppressed_download": f"/api/download/{supp_token}",
        "duplicates_download": f"/api/download/{dup_token}",
        "all_leads_download": f"/api/download/{all_token}",
        "duplicates_preview": logic.df_to_records(result["duplicates_df"], limit=200),
        "suppressed_preview": logic.df_to_records(suppressed_only_df, limit=preview_limit),
        "suppressed_preview_truncated": result["suppressed"] > preview_limit,
    })


# ---------------------------------------------------------------------------
# Tool 2: TAL : SUPP Scrubber
# ---------------------------------------------------------------------------

@app.post("/api/tal-scrubber")
async def tal_scrubber_endpoint(
    suppression_file: UploadFile = File(...),
    tal_file: UploadFile = File(...),
    out_filename: str = Form("scrubbed_tal"),
):
    supp_content, supp_name = await _read(suppression_file)
    tal_content, tal_name = await _read(tal_file)
    try:
        result = logic.tal_supp_scrubber(supp_content, supp_name, tal_content, tal_name)
    except ValueError as e:
        raise HTTPException(400, str(e))

    clean_bytes = logic.df_to_excel_bytes({"Clean TAL": result["clean_df"]})
    supp_bytes = logic.df_to_excel_bytes({"Suppressed": result["suppressed_df"]})
    clean_token = _store(f"{out_filename}_clean.xlsx", clean_bytes,
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    supp_token = _store(f"{out_filename}_suppressed.xlsx", supp_bytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return JSONResponse({
        "total_tal": result["total_tal"],
        "suppressed": result["suppressed"],
        "clean": result["clean"],
        "domain_column_used": result["domain_column_used"],
        "clean_download": f"/api/download/{clean_token}",
        "suppressed_download": f"/api/download/{supp_token}",
    })


# ---------------------------------------------------------------------------
# Tool 3: TAL : SUPP Overlap Analysis
# ---------------------------------------------------------------------------

@app.post("/api/tal-overlap")
async def tal_overlap_endpoint(
    tal_file: UploadFile = File(...),
    account_suppression_file: UploadFile = File(...),
    email_suppression_file: UploadFile = File(...),
    out_filename: str = Form("enriched_tal"),
):
    tal_content, tal_name = await _read(tal_file)
    acct_content, acct_name = await _read(account_suppression_file)
    email_content, email_name = await _read(email_suppression_file)
    try:
        result = logic.tal_overlap_analysis(
            tal_content, tal_name, acct_content, acct_name, email_content, email_name
        )
    except ValueError as e:
        raise HTTPException(400, str(e))

    bytes_out = logic.df_to_excel_bytes({"Enriched TAL": result["enriched_df"]})
    token = _store(f"{out_filename}.xlsx", bytes_out,
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return JSONResponse({
        "total_tal": result["total_tal"],
        "on_account_supp": result["on_account_supp"],
        "on_email_supp": result["on_email_supp"],
        "domain_column_used": result["domain_column_used"],
        "download": f"/api/download/{token}",
    })


# ---------------------------------------------------------------------------
# Tool 4a: List Combine
# ---------------------------------------------------------------------------

@app.post("/api/combine")
async def combine_endpoint(
    files: List[UploadFile] = File(...),
    out_filename: str = Form("combined"),
):
    pairs = []
    for f in files:
        content, name = await _read(f)
        pairs.append((content, name))
    try:
        result = logic.combine_lists(pairs)
    except ValueError as e:
        raise HTTPException(400, str(e))

    acct_bytes = logic.df_to_excel_bytes({"Accounts": result["account_df"]})
    email_bytes = logic.df_to_excel_bytes({"Emails": result["email_df"]})
    acct_token = _store(f"{out_filename}_accounts.xlsx", acct_bytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    email_token = _store(f"{out_filename}_emails.xlsx", email_bytes,
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return JSONResponse({
        "sources": result["sources"],
        "account_rows": result["account_rows"],
        "email_rows": result["email_rows"],
        "accounts_download": f"/api/download/{acct_token}",
        "emails_download": f"/api/download/{email_token}",
    })


# ---------------------------------------------------------------------------
# Tool 4b: List Split
# ---------------------------------------------------------------------------

@app.post("/api/split/columns")
async def split_columns_endpoint(file: UploadFile = File(...)):
    content, name = await _read(file)
    try:
        df = logic.read_single_table(content, name)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return JSONResponse({"columns": list(df.columns)})


@app.post("/api/split")
async def split_endpoint(
    file: UploadFile = File(...),
    column: str = Form(...),
    out_filename: str = Form("split_output"),
):
    content, name = await _read(file)
    try:
        result = logic.split_list(content, name, column)
    except ValueError as e:
        raise HTTPException(400, str(e))

    token = _store(f"{out_filename}.zip", result["zip_bytes"], "application/zip")
    return JSONResponse({
        "groups": result["groups"],
        "download": f"/api/download/{token}",
    })


# ---------------------------------------------------------------------------
# Tool 5: List Comparison
# ---------------------------------------------------------------------------

@app.post("/api/list-compare")
async def list_compare_endpoint(
    list1: Optional[UploadFile] = File(None),
    list2: Optional[UploadFile] = File(None),
    list3: Optional[UploadFile] = File(None),
    list1_text: Optional[str] = Form(None),
    list2_text: Optional[str] = Form(None),
    list3_text: Optional[str] = Form(None),
    dedupe_within: bool = Form(True),
    out_filename: str = Form("list_comparison"),
):
    pairs = []
    for idx, (f, t) in enumerate([(list1, list1_text), (list2, list2_text), (list3, list3_text)], start=1):
        has_file = f is not None and getattr(f, "filename", None)
        has_text = t and t.strip()
        if not has_file and not has_text:
            if idx <= 2:
                raise HTTPException(400, f"List {idx} is required (upload a file or paste values).")
            continue
        content, name = await _resolve_input(f, t, f"list{idx}")
        pairs.append((content, name))

    try:
        result = logic.compare_lists(pairs, dedupe_within=dedupe_within)
    except ValueError as e:
        raise HTTPException(400, str(e))
    bytes_out = logic.df_to_excel_bytes({"Comparison": result["result_df"]})
    token = _store(f"{out_filename}.xlsx", bytes_out,
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return JSONResponse({
        "total_unique": result["total_unique"],
        "in_all": result["in_all"],
        "in_exactly_2": result["in_exactly_2"],
        "in_exactly_1": result["in_exactly_1"],
        "per_list_counts": result["per_list_counts"],
        "download": f"/api/download/{token}",
    })


# ---------------------------------------------------------------------------
# Frontend
# ---------------------------------------------------------------------------

app.mount("/", StaticFiles(directory="static", html=True), name="static")
