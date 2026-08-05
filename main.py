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

Uploads are parsed and transformed in-memory. Generated result files are
written to a temp directory on disk (not held as bytes in-process) and are
served for download until they expire from a short-lived TTL cache.
"""

import asyncio
import gc
import os
import tempfile
import time
import uuid
from typing import List, Optional

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import logic

app = FastAPI(title="Suppression & TAL Toolkit")

# Reject uploads above this size outright, before they're fully read into
# memory, instead of letting a huge file ride all the way through parsing
# and Excel-writing and OOM-kill the process partway. Default matches the
# actual ceiling this deployment (Render Free, 512MB RAM) can process
# without more RAM: a 40MB upload is roughly the largest this pipeline can
# turn into a pandas DataFrame + Excel output(s) and stay under 512MB.
# Override with the MAX_UPLOAD_MB env var — but raising it does NOT make
# bigger files safe on this instance; it only changes where the line is
# drawn. See README for the underlying memory math.
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_MB", "40")) * 1024 * 1024

# At most this many uploads are parsed/transformed at once. Each concurrent
# request multiplies peak memory (raw bytes + DataFrame(s) + Excel output),
# so on a memory-constrained deployment two overlapping 40MB requests can
# OOM where one at a time would not. Extra requests wait for a slot instead
# of running in parallel — same inputs, same outputs, just serialized.
# Override with MAX_CONCURRENT_JOBS if the deployment ever gets more RAM.
_PROCESSING_SEMAPHORE = asyncio.Semaphore(int(os.environ.get("MAX_CONCURRENT_JOBS", "1")))

# Generated result files are written to disk (an ephemeral temp dir), not
# kept as bytes in a process-wide dict. Holding N large xlsx blobs in RAM
# for the full TTL — on top of whatever a concurrent request is parsing —
# was the other big memory multiplier; only file paths + metadata live in
# memory now. {token: {"filename": str, "path": str, "media_type": str, "created": float}}
_RESULTS = {}
_RESULT_TTL_SECONDS = 20 * 60
_RESULTS_DIR = tempfile.mkdtemp(prefix="toolkit_results_")


def _purge_expired_results():
    now = time.time()
    expired = [t for t, item in _RESULTS.items() if now - item["created"] > _RESULT_TTL_SECONDS]
    for t in expired:
        item = _RESULTS.pop(t)
        try:
            os.remove(item["path"])
        except OSError:
            pass


def _store(filename: str, data: bytes, media_type: str) -> str:
    _purge_expired_results()
    token = uuid.uuid4().hex
    path = os.path.join(_RESULTS_DIR, token)
    with open(path, "wb") as f:
        f.write(data)
    _RESULTS[token] = {"filename": filename, "path": path, "media_type": media_type, "created": time.time()}
    return token


@app.get("/api/download/{token}")
def download(token: str):
    item = _RESULTS.get(token)
    if not item or not os.path.exists(item["path"]):
        raise HTTPException(404, "Result not found or expired. Re-run the tool.")
    return FileResponse(item["path"], media_type=item["media_type"], filename=item["filename"])


async def _read(upload: UploadFile):
    # Read in bounded chunks and bail out the moment the cap is crossed,
    # instead of buffering an unlimited amount before ever checking size.
    # Accumulating into a bytearray (grown in place) instead of a list of
    # chunks joined at the end avoids briefly holding the data twice —
    # list-of-chunks + b"".join(chunks) has both the chunk list and the
    # freshly-joined bytes alive at once for the moment join() runs.
    buf = bytearray()
    while True:
        chunk = await upload.read(1024 * 1024)
        if not chunk:
            break
        buf.extend(chunk)
        if len(buf) > MAX_UPLOAD_BYTES:
            raise HTTPException(
                413,
                f"{upload.filename} exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)}MB "
                f"upload limit for this deployment.",
            )
    content = bytes(buf)
    del buf
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
    async with _PROCESSING_SEMAPHORE:
        content, name = await _resolve_input(file, text, "leads_file")
        try:
            df = logic.read_single_table(content, name)
        except ValueError as e:
            raise HTTPException(400, str(e))
        del content
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
    async with _PROCESSING_SEMAPHORE:
        supp_content, supp_name = await _resolve_input(suppression_file, suppression_text, "suppression_file")
        lead_content, lead_name = await _resolve_input(leads_file, leads_text, "leads_file")
        try:
            result = logic.email_scrubber(supp_content, supp_name, lead_content, lead_name, email_column)
        except ValueError as e:
            raise HTTPException(400, str(e))
        # Both are fully parsed into result's DataFrames by this point and
        # aren't touched again, but as local variables they'd otherwise stay
        # alive for the rest of this (fairly long) function.
        del supp_content, lead_content

        annotated_df = result["annotated_df"]
        mask = result["suppression_mask"]
        preview_limit = 500

        # Build + store each output one at a time, dropping each DataFrame as soon
        # as it's serialized, instead of materializing clean/suppressed/duplicates
        # all at once alongside annotated_df. That "all copies alive together"
        # pattern was the biggest avoidable memory multiplier here — up to 4 full
        # copies of the list at peak — and a likely contributor to OOMs on
        # Render's 512MB instance even on lists well under 100MB.
        all_bytes = logic.df_to_excel_bytes({"All Leads": annotated_df})
        all_token = _store(f"{out_filename}_all_leads.xlsx", all_bytes,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        del all_bytes

        suppressed_df = annotated_df[mask]
        suppressed_preview = logic.df_to_records(suppressed_df.head(preview_limit), limit=preview_limit)
        supp_bytes = logic.df_to_excel_bytes({"Suppressed": suppressed_df})
        del suppressed_df
        supp_token = _store(f"{out_filename}_suppressed.xlsx", supp_bytes,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        del supp_bytes

        clean_df = annotated_df[~mask]
        clean_bytes = logic.df_to_excel_bytes({"Clean Leads": clean_df})
        del clean_df
        clean_token = _store(f"{out_filename}_clean.xlsx", clean_bytes,
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        del clean_bytes

        duplicates_df = result["duplicates_df"]
        dup_bytes = logic.df_to_excel_bytes({"Duplicates": duplicates_df})
        dup_token = _store(f"{out_filename}_duplicates.xlsx", dup_bytes,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        del dup_bytes

        response = JSONResponse({
            "total_leads": result["total_leads"],
            "suppressed": result["suppressed"],
            "clean": result["clean"],
            "duplicates": result["duplicates"],
            "email_column_used": result["email_column_used"],
            "clean_download": f"/api/download/{clean_token}",
            "suppressed_download": f"/api/download/{supp_token}",
            "duplicates_download": f"/api/download/{dup_token}",
            "all_leads_download": f"/api/download/{all_token}",
            "duplicates_preview": logic.df_to_records(duplicates_df, limit=200),
            "suppressed_preview": suppressed_preview,
            "suppressed_preview_truncated": result["suppressed"] > preview_limit,
        })
        # This endpoint is the heaviest in the app (4 full Excel builds off
        # one lead list). df_to_excel_bytes() already collects after each
        # xlsxwriter workbook it builds, but annotated_df/result themselves
        # are about to go out of scope here too — one last sweep before the
        # semaphore releases the next queued request means this request's
        # garbage doesn't linger into the next one's peak. Note this frees
        # Python-heap garbage; it does not force the OS to reclaim RSS, but
        # it does keep that memory available for reuse within this process.
        del result, annotated_df, mask, duplicates_df
        gc.collect()
        return response


# ---------------------------------------------------------------------------
# Tool 2: TAL : SUPP Scrubber
# ---------------------------------------------------------------------------

@app.post("/api/tal-scrubber")
async def tal_scrubber_endpoint(
    suppression_file: UploadFile = File(...),
    tal_file: UploadFile = File(...),
    out_filename: str = Form("scrubbed_tal"),
):
    async with _PROCESSING_SEMAPHORE:
        supp_content, supp_name = await _read(suppression_file)
        tal_content, tal_name = await _read(tal_file)
        try:
            result = logic.tal_supp_scrubber(supp_content, supp_name, tal_content, tal_name)
        except ValueError as e:
            raise HTTPException(400, str(e))
        del supp_content, tal_content

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
    async with _PROCESSING_SEMAPHORE:
        tal_content, tal_name = await _read(tal_file)
        acct_content, acct_name = await _read(account_suppression_file)
        email_content, email_name = await _read(email_suppression_file)
        try:
            result = logic.tal_overlap_analysis(
                tal_content, tal_name, acct_content, acct_name, email_content, email_name
            )
        except ValueError as e:
            raise HTTPException(400, str(e))
        del tal_content, acct_content, email_content

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
    async with _PROCESSING_SEMAPHORE:
        pairs = []
        for f in files:
            content, name = await _read(f)
            pairs.append((content, name))
        try:
            result = logic.combine_lists(pairs)
        except ValueError as e:
            raise HTTPException(400, str(e))
        del pairs

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
    async with _PROCESSING_SEMAPHORE:
        content, name = await _read(file)
        try:
            df = logic.read_single_table(content, name)
        except ValueError as e:
            raise HTTPException(400, str(e))
        del content
        return JSONResponse({"columns": list(df.columns)})


@app.post("/api/split")
async def split_endpoint(
    file: UploadFile = File(...),
    column: str = Form(...),
    out_filename: str = Form("split_output"),
):
    async with _PROCESSING_SEMAPHORE:
        content, name = await _read(file)
        try:
            result = logic.split_list(content, name, column)
        except ValueError as e:
            raise HTTPException(400, str(e))
        del content

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
    async with _PROCESSING_SEMAPHORE:
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
        del pairs

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
