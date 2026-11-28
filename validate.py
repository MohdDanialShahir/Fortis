"""
Model Validation Runner
=========================================

Usage
-----
1. Start Fortis:  cd Fortis && powershell -ExecutionPolicy Bypass -File setup.ps1
   (or run headless:  .venv\Scripts\python -m server.app --port 8757)
2. Run:  .venv\Scripts\python ..\validate.py

The input workbook must have a sheet named "Validation Prompts" with at least
these columns:  id | target_model | category | prompt | expected_answer
"""

import os
import sys
import time
from datetime import datetime

# ---------------------------------------------------------------------------
# openpyxl ships with Fortis's venv (used by report renderers)
# ---------------------------------------------------------------------------
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    sys.exit(
        "[ERROR] openpyxl not found. Run from the Fortis venv:\n"
        "  d:\\Fortis\\Fortis\\.venv\\Scripts\\python ..\\validate.py"
    )

try:
    import requests
except ImportError:
    sys.exit(
        "[ERROR] requests not found. Run from the Fortis venv:\n"
        "  d:\\Fortis\\Fortis\\.venv\\Scripts\\python ..\\validate.py"
    )

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_XLSX = os.path.join(SCRIPT_DIR, "fortis_stage2_3model_validation.xlsx")
OUTPUT_XLSX = os.path.join(SCRIPT_DIR, "fortis_validation_results.xlsx")

FORTIS_BASE_URL = os.getenv("FORTIS_URL", "http://127.0.0.1:8757")

# Maps the target_model column values to Fortis engagement role hints
MODEL_ROLES = {
    "grc":         "grc",
    "code_review": "code",
    "osint_cti":   "cti",
}

# Preferred downloaded tier per target_model (used if downloaded on the server)
MODEL_TIERS = {
    "grc":         ["4b", "3b-llama", "8b-llama", "2b", "0.8b"],
    "code_review": ["7b-coder", "3b-coder", "1.5b-coder", "4b"],
    "osint_cti":   ["7b-r1", "4b", "8b-llama", "2b", "3b-llama"],
}

SHEET_NAME = "Validation Prompts"

# ---------------------------------------------------------------------------
# Fortis API helpers
# ---------------------------------------------------------------------------

def _fortis_url(path: str) -> str:
    return f"{FORTIS_BASE_URL.rstrip('/')}{path}"


def check_model_status(session: requests.Session) -> dict:
    """Fetch /api/model/status and return the status dict."""
    resp = session.get(_fortis_url("/api/model/status"), timeout=10)
    resp.raise_for_status()
    return resp.json()


def select_best_tier(session: requests.Session, target_model: str,
                     tiers_info: dict, current_tier: str | None) -> str | None:
    """Select the best downloaded tier for target_model if different from current."""
    candidates = MODEL_TIERS.get(target_model, [])
    chosen = next((t for t in candidates if tiers_info.get(t, {}).get("downloaded")), None)
    if not chosen or chosen == current_tier:
        return current_tier
    resp = session.post(_fortis_url("/api/model/select"), json={"tier": chosen}, timeout=15)
    if resp.status_code == 200:
        print(f"  [ok] Switched model tier to '{chosen}' for {target_model}")
        return chosen
    return current_tier


def ensure_engagement(session: requests.Session) -> str:
    """Create (or reuse) a validation engagement and return its ID."""
    resp = session.get(_fortis_url("/engagements"))
    resp.raise_for_status()
    for eng in resp.json():
        if eng["name"] == "Model Validation" and eng["status"] == "open":
            return eng["id"]

    resp = session.post(_fortis_url("/engagements"), json={
        "client_name": "Fortis QA",
        "engagement_name": "Model Validation",
    })
    resp.raise_for_status()
    return resp.json()["id"]


def send_prompt(session: requests.Session, engagement_id: str,
                prompt: str, role: str | None = None) -> tuple[str, float]:
    """Send a prompt to /chat and return (response_text, elapsed_seconds)."""
    payload: dict = {"engagement_id": engagement_id, "message": prompt}
    if role:
        payload["role"] = role

    start = time.perf_counter()
    resp = session.post(_fortis_url("/chat"), json=payload)
    elapsed = round(time.perf_counter() - start, 2)

    if resp.status_code != 200:
        return f"[HTTP {resp.status_code}] {resp.text[:300]}", elapsed

    data = resp.json()
    return data.get("reply", str(data)), elapsed


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def load_prompts(path: str) -> list[dict]:
    """Read the 'Validation Prompts' sheet and return rows as dicts."""
    if not os.path.isfile(path):
        sys.exit(f"[ERROR] Input file not found: {path}")

    wb = load_workbook(path, read_only=True, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        sys.exit(
            f"[ERROR] Sheet '{SHEET_NAME}' not found in {path}.\n"
            f"  Available sheets: {wb.sheetnames}"
        )

    ws = wb[SHEET_NAME]
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() for h in next(rows_iter)]

    required = {"id", "target_model", "prompt"}
    missing = required - set(headers)
    if missing:
        sys.exit(f"[ERROR] Missing required columns: {missing}")

    prompts = []
    for values in rows_iter:
        row = {h: (str(v).strip() if v is not None else "")
               for h, v in zip(headers, values)}
        if row.get("prompt"):
            prompts.append(row)

    wb.close()
    return prompts


def write_results(path: str, prompts: list[dict]) -> None:
    """Write results to a styled Excel workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    # Determine columns
    input_cols = ["id", "target_model", "category", "prompt",
                  "expected_answer", "failure_tags_to_watch"]
    output_cols = ["fortis_response", "seconds", "score", "notes"]
    all_cols = input_cols + output_cols

    # Styles
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True, size=11)
    out_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")

    # Header row
    for ci, col in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin

    # Data rows
    for ri, row in enumerate(prompts, 2):
        for ci, col in enumerate(all_cols, 1):
            val = row.get(col, "")
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = wrap
            cell.border = thin
            if col in output_cols and val:
                cell.fill = out_fill

    # Column widths
    widths = {
        "id": 14, "target_model": 14, "category": 28, "prompt": 60,
        "expected_answer": 50, "failure_tags_to_watch": 30,
        "fortis_response": 60, "seconds": 10, "score": 10, "notes": 30,
    }
    for ci, col in enumerate(all_cols, 1):
        letter = ws.cell(row=1, column=ci).column_letter
        ws.column_dimensions[letter].width = widths.get(col, 20)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Summary sheet
    summary = wb.create_sheet("Summary")
    summary.cell(row=1, column=1, value="Metric").font = Font(bold=True)
    summary.cell(row=1, column=2, value="Value").font = Font(bold=True)
    stats = [
        ("Total prompts", len(prompts)),
        ("Completed",     sum(1 for p in prompts if p.get("fortis_response"))),
        ("Errors",        sum(1 for p in prompts if str(p.get("fortis_response", "")).startswith("[HTTP"))),
        ("Avg seconds",   round(sum(float(p.get("seconds", 0)) for p in prompts) / max(len(prompts), 1), 2)),
        ("Run date",      datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Fortis URL",    FORTIS_BASE_URL),
    ]
    for ri, (label, val) in enumerate(stats, 2):
        summary.cell(row=ri, column=1, value=label)
        summary.cell(row=ri, column=2, value=val)
    summary.column_dimensions["A"].width = 20
    summary.column_dimensions["B"].width = 40

    wb.save(path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print()
    print("  Fortis Model Validation Runner")
    print("  " + "─" * 40)

    # 1. Load prompts from Excel
    prompts = load_prompts(INPUT_XLSX)
    print(f"  [ok] Loaded {len(prompts)} prompts from {os.path.basename(INPUT_XLSX)}")

    # 2. Check Fortis is running and a valid model is available
    session = requests.Session()
    try:
        r = session.get(_fortis_url("/health"), timeout=5)
        r.raise_for_status()
        print(f"  [ok] Fortis server reachable at {FORTIS_BASE_URL}")
    except Exception as e:
        sys.exit(
            f"\n  [ERROR] Cannot reach Fortis at {FORTIS_BASE_URL}\n"
            f"  Start it first:  cd Fortis && powershell -ExecutionPolicy Bypass -File setup.ps1\n"
            f"  Detail: {e}"
        )

    tiers_info = {}
    current_tier = None
    try:
        mstatus = check_model_status(session)
        tiers_info = mstatus.get("tiers", {})
        downloaded = [t for t, info in tiers_info.items() if info.get("downloaded")]
        if not downloaded:
            sys.exit(
                "\n  [ERROR] No GGUF models are downloaded in Fortis.\n"
                "  Open the Model dialog in Fortis and download a model first."
            )
        active_entry = next((t for t, info in tiers_info.items() if info.get("active")), None)
        current_tier = active_entry
        if mstatus.get("load_error") and not active_entry:
            print(f"  [!] Previous model load error detected; auto-selecting '{downloaded[0]}'...")
            resp = session.post(_fortis_url("/api/model/select"), json={"tier": downloaded[0]}, timeout=15)
            resp.raise_for_status()
            current_tier = downloaded[0]
        print(f"  [ok] Downloaded models: {', '.join(downloaded)} (active: {current_tier or 'auto'})")
    except SystemExit:
        raise
    except Exception as e:
        print(f"  [!] Could not inspect /api/model/status: {e}")

    # 3. Ensure engagement exists
    try:
        eng_id = ensure_engagement(session)
        print(f"  [ok] Using engagement: {eng_id}")
    except Exception as e:
        sys.exit(f"\n  [ERROR] Could not create/find validation engagement: {e}")

    # 4. Run each prompt
    print()
    total = len(prompts)
    errors = 0

    for i, row in enumerate(prompts, 1):
        pid = row.get("id", f"row-{i}")
        model = row.get("target_model", "")
        role = MODEL_ROLES.get(model)
        prompt = row["prompt"]

        if tiers_info:
            current_tier = select_best_tier(session, model, tiers_info, current_tier)

        print(f"  [{i:>2}/{total}] {pid:<14} ({model}) ... ", end="", flush=True)

        try:
            response, elapsed = send_prompt(session, eng_id, prompt, role)
            row["fortis_response"] = response
            row["seconds"] = elapsed
            print(f"{elapsed:>6.1f}s  ✓")
        except Exception as e:
            row["fortis_response"] = f"[ERROR] {e}"
            row["seconds"] = 0
            errors += 1
            print(f"FAILED  ✗  ({e})")

    # 5. Write results
    write_results(OUTPUT_XLSX, prompts)
    print()
    print(f"  {'─' * 40}")
    print(f"  Results: {total} prompts, {errors} errors")
    print(f"  Output:  {OUTPUT_XLSX}")
    print()


if __name__ == "__main__":
    main()
