"""
generate_report.py
Generates the Approval Summary report from the input approval file.

Changes:
  1. Pivot (Sheet1) now includes Pending column
  2. Sheet2 (rejection reasons) preserved as-is
  3. All Approval Summary: duplicate ApprovalNum rows are highlighted in yellow
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ── Column definitions ────────────────────────────────────────────────────────

TABLE_COLS = [
    "Table1[CoCode]", "Table1[CoName]", "Table1[PatientFileNum]",
    "Table1[PatientName]", "Table1[DoctorCode]", "Table1[DoctorName]",
    "Table1[ServiceCode]", "Table1[ServiceName]", "Table1[ApprovalNum]",
    "Table1[Quantity]", "Table1[ApprovedStatus]", "Table1[ApprovedQuantity]",
    "Table1[ApprovalDate]", "Table1[ApprovalDtlID]", "Table1[RequestDateTime]",
    "Table1[ApprovalSentDateandTime]", "Table1[ApprovalResponceDateAndTime]",
    "Table1[Time1Days]", "Table1[Time1Hours]", "Table1[Time1Minutes]",
    "Table1[Time2Days]", "Table1[Time2Hours]", "Table1[Time2Minutes]",
    "Table1[ResponseRemarks]",
]

SOURCE_COLS = [
    "CoCode", "CoName", "PatientFileNum", "PatientName", "DoctorCode",
    "DoctorName", "ServiceCode", "ServiceName", "ApprovalNum", "Quantity",
    "ApprovedStatus", "ApprovedQuantity", "ApprovalDate", "ApprovalDtlID",
    "RequestDateTime", "ApprovalSentDateandTime", "ApprovalResponceDateAndTime",
    "Time1Days", "Time1Hours", "Time1Minutes", "Time2Days", "Time2Hours",
    "Time2Minutes", "ResponseRemarks",
]

REQUIRED_COLS = ["CoCode", "CoName", "ApprovedStatus", "ApprovalNum", "ResponseRemarks"]

# Highlight colours
YELLOW = PatternFill("solid", fgColor="FFFF00")   # duplicate ApprovalNum
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, color="FFFFFF")


# ── Sheet detection ───────────────────────────────────────────────────────────

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
            f"None of the sheets contain the expected columns. "
            f"Sheets found: {xl.sheet_names}. "
            f"Expected columns like: {REQUIRED_COLS}"
        )
    return best_name, best_df


# ── Helpers ───────────────────────────────────────────────────────────────────

def _auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 50)


def _write_cell(cell, val):
    if isinstance(val, float) and pd.isna(val):
        cell.value = None
    elif hasattr(val, 'item'):
        cell.value = val.item()
    else:
        cell.value = val


# ── Sheet writers ─────────────────────────────────────────────────────────────

def _write_detail_sheet(ws, title, df):
    """Sheet4 / Sheet5: title row1, blank row2, Table1[x] headers row3, data row4+"""
    ws.cell(row=1, column=1, value=title).font = Font(bold=True)
    for ci, col_name in enumerate(TABLE_COLS, 1):
        ws.cell(row=3, column=ci, value=col_name)
    for ri, row in enumerate(df.itertuples(index=False), 4):
        for ci, val in enumerate(row, 1):
            _write_cell(ws.cell(row=ri, column=ci), val)
    _auto_width(ws)


def _write_summary_sheet(ws, df):
    """
    All Approval Summary:
      Row 1 = plain column headers
      Row 2+ = data
      Duplicate ApprovalNum rows highlighted yellow
      ApprovalNum column header bold-noted
    """
    # Find duplicate ApprovalNums
    approval_col_idx = SOURCE_COLS.index("ApprovalNum")  # 0-based → col 9 (1-based)
    dup_mask = df.duplicated("ApprovalNum", keep=False).tolist()

    # Row 1 – headers
    for ci, col_name in enumerate(SOURCE_COLS, 1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    # Mark ApprovalNum header to flag duplicates
    appr_header_cell = ws.cell(row=1, column=approval_col_idx + 1)
    appr_header_cell.value = "ApprovalNum ⚠ (yellow = duplicate)"

    # Row 2+ – data with conditional highlighting
    for ri, (row, is_dup) in enumerate(
        zip(df.itertuples(index=False), dup_mask), 2
    ):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci)
            _write_cell(cell, val)
            if is_dup:
                cell.fill = YELLOW

    _auto_width(ws)


def _write_sheet1(ws, df):
    """
    Pivot: Distinct Count of ApprovalNum by CoName x ApprovedStatus
    Includes: Approved | Pending | Partiaily | Rejected | Grand Total
    """
    pivot = (
        df.groupby(["CoName", "ApprovedStatus"])["ApprovalNum"]
        .nunique()
        .unstack(fill_value=0)
    )
    # Ensure all expected status columns exist
    for col in ["Approved", "Pending", "Partiaily", "Rejected"]:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot[["Approved", "Pending", "Partiaily", "Rejected"]]
    pivot["Grand Total"] = pivot.sum(axis=1)
    grand = pivot.sum(axis=0).rename("Grand Total")
    pivot = pd.concat([pivot, grand.to_frame().T])
    pivot.index.name = "Row Labels"
    pivot = pivot.reset_index()

    # Row 3 – label row
    ws.cell(row=3, column=1, value="Distinct Count of ApprovalNum")
    ws.cell(row=3, column=2, value="Column Labels")

    # Row 4 – column headers
    for ci, col in enumerate(pivot.columns, 1):
        cell = ws.cell(row=4, column=ci, value=col)
        cell.font = Font(bold=True)

    # Row 5+ – data
    for ri, row in enumerate(pivot.itertuples(index=False), 5):
        for ci, val in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=val.item() if hasattr(val, 'item') else val)

    _auto_width(ws)


def _write_sheet2(ws, df):
    """
    Rejection reasons count
    Row 1: ApprovedStatus | Rejected
    Row 3: Row Labels | Count of ResponseRemarks
    Row 4+: data + Grand Total
    """
    rejected = df[df["ApprovedStatus"] == "Rejected"]
    reasons  = rejected["ResponseRemarks"].value_counts().reset_index()
    reasons.columns = ["Row Labels", "Count of ResponseRemarks"]
    grand = pd.DataFrame(
        [["Grand Total", int(reasons["Count of ResponseRemarks"].sum())]],
        columns=reasons.columns,
    )
    reasons = pd.concat([reasons, grand], ignore_index=True)

    ws.cell(row=1, column=1, value="ApprovedStatus")
    ws.cell(row=1, column=2, value="Rejected")
    ws.cell(row=3, column=1, value="Row Labels")
    ws.cell(row=3, column=2, value="Count of ResponseRemarks")
    for ri, row in enumerate(reasons.itertuples(index=False), 4):
        ws.cell(row=ri, column=1, value=row[0])
        ws.cell(row=ri, column=2, value=row[1])
    ws.column_dimensions["A"].width = 120
    ws.column_dimensions["B"].width = 25


# ── Public API ────────────────────────────────────────────────────────────────

def generate_report(approval_file, output_path):
    xl = pd.ExcelFile(approval_file)
    detected_sheet, df = _find_sheet(xl)
    df.columns = df.columns.str.strip()

    for col in SOURCE_COLS:
        if col not in df.columns:
            df[col] = None
    df = df[SOURCE_COLS]

    bupa    = df[df["CoName"].str.contains("Bupa", case=False, na=False)].copy()
    partial = df[df["ApprovedStatus"] == "Partiaily"].copy()
    pending = df[df["ApprovedStatus"] == "Pending"].copy()

    # Duplicate ApprovalNum stats
    dup_count = int(df.duplicated("ApprovalNum", keep=False).sum())

    wb = Workbook()
    wb.remove(wb.active)

    ws4 = wb.create_sheet("Sheet4")
    _write_detail_sheet(ws4,
        "Data returned for Distinct Count of ApprovalNum, "
        "Bupa Arabia for Cooperative Insurance (First 1000 rows).", bupa)

    ws5 = wb.create_sheet("Sheet5")
    _write_detail_sheet(ws5,
        "Data returned for Distinct Count of ApprovalNum, "
        "Partiaily (First 1000 rows).", partial)

    ws1 = wb.create_sheet("Sheet1")
    _write_sheet1(ws1, df)

    ws2 = wb.create_sheet("Sheet2")
    _write_sheet2(ws2, df)

    ws_all = wb.create_sheet("All Approval Summary")
    _write_summary_sheet(ws_all, df)

    wb.save(output_path)

    return {
        "total":          len(df),
        "approved":       len(df[df["ApprovedStatus"] == "Approved"]),
        "partial":        len(partial),
        "pending":        len(pending),
        "rejected":       len(df[df["ApprovedStatus"] == "Rejected"]),
        "bupa":           len(bupa),
        "duplicates":     dup_count,
        "detected_sheet": detected_sheet,
    }
