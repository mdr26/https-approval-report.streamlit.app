"""
generate_report.py
Deduplication rule:
  - Keep one row per unique (ApprovalNum, ApprovedStatus) combination
  - e.g. 3x Approved + 1x Rejected → 1x Approved + 1x Rejected (2 rows)
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SOURCE_COLS = [
    "CoCode", "CoName", "PatientFileNum", "PatientName", "DoctorCode",
    "DoctorName", "ServiceCode", "ServiceName", "ApprovalNum", "Quantity",
    "ApprovedStatus", "ApprovedQuantity", "ApprovalDate", "ApprovalDtlID",
    "RequestDateTime", "ApprovalSentDateandTime", "ApprovalResponceDateAndTime",
    "Time1Days", "Time1Hours", "Time1Minutes", "Time2Days", "Time2Hours",
    "Time2Minutes", "ResponseRemarks",
]

REQUIRED_COLS = ["CoCode", "CoName", "ApprovedStatus", "ApprovalNum", "ResponseRemarks"]

HEADER_FILL = PatternFill("solid", fgColor="1a2e44")
HEADER_FONT = Font(bold=True, color="FFFFFF")
THIN        = Side(style="thin")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _find_sheet(xl):
    best_name, best_df, best_hits = None, None, -1
    for name in xl.sheet_names:
        try:
            df = xl.parse(name)
            df.columns = df.columns.str.strip()
            hits = sum(1 for c in REQUIRED_COLS if c in df.columns)
            if hits > best_hits:
                best_hits, best_name, best_df = hits, name, df
            if hits == len(REQUIRED_COLS):
                break
        except Exception:
            continue
    if best_df is None or best_hits == 0:
        raise ValueError(
            f"Could not find the expected columns in any sheet.\n"
            f"Sheets found: {xl.sheet_names}\n"
            f"Expected columns: {REQUIRED_COLS}"
        )
    return best_name, best_df


def _auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 55)


def generate_report(approval_file, output_path):
    # ── Load & detect sheet ───────────────────────────────────────────────
    xl = pd.ExcelFile(approval_file)
    detected_sheet, df = _find_sheet(xl)
    df.columns = df.columns.str.strip()

    # Ensure all expected columns exist
    for col in SOURCE_COLS:
        if col not in df.columns:
            df[col] = None
    df = df[SOURCE_COLS].copy()

    rows_before = len(df)

    # ── Deduplicate: keep first row per (ApprovalNum, ApprovedStatus) ─────
    # This means: if 3 Approved rows for same ApprovalNum → keep 1 Approved
    # But if 1 Approved + 1 Rejected for same ApprovalNum → keep both
    df = df.drop_duplicates(subset=["ApprovalNum", "ApprovedStatus"], keep="first")
    df = df.reset_index(drop=True)

    rows_after  = len(df)
    rows_removed = rows_before - rows_after

    # ── Build workbook ────────────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "All Approval Summary"

    # Row 1 — headers
    for ci, col_name in enumerate(SOURCE_COLS, 1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[1].height = 20

    # Row 2+ — data
    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci)
            if isinstance(val, float) and pd.isna(val):
                cell.value = None
            elif hasattr(val, 'item'):
                cell.value = val.item()
            else:
                cell.value = val
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    _auto_width(ws)

    wb.save(output_path)

    return {
        "total":          len(df),
        "approved":       int((df["ApprovedStatus"] == "Approved").sum()),
        "partial":        int((df["ApprovedStatus"] == "Partiaily").sum()),
        "pending":        int((df["ApprovedStatus"] == "Pending").sum()),
        "rejected":       int((df["ApprovedStatus"] == "Rejected").sum()),
        "rows_removed":   rows_removed,
        "detected_sheet": detected_sheet,
    }
