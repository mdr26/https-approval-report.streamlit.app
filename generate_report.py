"""
generate_report.py
Generates a single-sheet report: All Approval Summary
- Exact 24 columns, plain headers row 1, data from row 2
- Duplicate ApprovalNum rows highlighted yellow
- No other sheets
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

YELLOW      = PatternFill("solid", fgColor="FFFF00")
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

    # ── Find duplicates ───────────────────────────────────────────────────
    dup_mask = df.duplicated("ApprovalNum", keep=False).tolist()
    dup_count = sum(dup_mask)

    # ── Build workbook — single sheet ─────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "All Approval Summary"

    # Row 1 — headers
    for ci, col_name in enumerate(SOURCE_COLS, 1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.border = BORDER
    ws.row_dimensions[1].height = 20

    # Row 2+ — data
    for ri, (row, is_dup) in enumerate(zip(df.itertuples(index=False), dup_mask), 2):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci)
            # Clean value
            if isinstance(val, float) and pd.isna(val):
                cell.value = None
            elif hasattr(val, 'item'):
                cell.value = val.item()
            else:
                cell.value = val
            # Highlight duplicates
            if is_dup:
                cell.fill = YELLOW
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")

    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-fit column widths
    _auto_width(ws)

    wb.save(output_path)

    return {
        "total":          len(df),
        "approved":       int((df["ApprovedStatus"] == "Approved").sum()),
        "partial":        int((df["ApprovedStatus"] == "Partiaily").sum()),
        "pending":        int((df["ApprovedStatus"] == "Pending").sum()),
        "rejected":       int((df["ApprovedStatus"] == "Rejected").sum()),
        "duplicates":     dup_count,
        "detected_sheet": detected_sheet,
    }
