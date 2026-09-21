#!/usr/bin/env python3
"""
High-Performance Excel MCP Server powered by Python, Pandas, OpenPyXL,
and Native Windows Microsoft Excel (COM Automation via PyWin32).
"""

import os
import math
from typing import Any, Optional, List, Dict, Union
import pandas as pd
import openpyxl
from mcp.server.fastmcp import FastMCP

# Try importing pywin32 for native Windows Excel COM automation
HAS_WIN32 = False
try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

mcp = FastMCP("excel-tools")


# ==========================================
# 1. PANDAS & OPENPYXL FAST DATA TOOLS
# ==========================================

@mcp.tool()
def get_workbook_info(file_path: str) -> Dict[str, Any]:
    """
    Get metadata for an Excel workbook without loading entire sheets into RAM.
    Returns sheet names, dimensions, column headers preview, and file size in MB.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    file_size_mb = round(os.path.getsize(file_path) / (1024 * 1024), 2)
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets_info = []

    try:
        for name in wb.sheetnames:
            ws = wb[name]
            headers = []
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                headers = [str(cell) if cell is not None else f"Column_{i+1}" for i, cell in enumerate(row)]
                break

            sheets_info.append({
                "name": name,
                "max_rows": ws.max_row,
                "max_columns": ws.max_column,
                "header_count": len(headers),
                "headers_preview": headers[:30],
            })
    finally:
        wb.close()

    return {
        "file_path": file_path,
        "file_size_mb": file_size_mb,
        "total_sheets": len(sheets_info),
        "sheets": sheets_info,
    }


@mcp.tool()
def preview_sheet(
    file_path: str,
    sheet_name: Optional[str] = None,
    nrows: int = 10,
) -> Dict[str, Any]:
    """
    Preview the first N rows of a sheet. Loads only the requested rows for speed.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    nrows = max(1, min(nrows, 200))
    df = pd.read_excel(file_path, sheet_name=sheet_name or 0, nrows=nrows)
    clean_df = df.where(pd.notnull(df), None)

    return {
        "sheet_name": sheet_name or "First Sheet",
        "preview_row_count": len(clean_df),
        "columns": list(df.columns),
        "column_types": {col: str(dtype) for col, dtype in df.dtypes.items()},
        "rows": clean_df.to_dict(orient="records"),
    }


@mcp.tool()
def query_rows(
    file_path: str,
    sheet_name: Optional[str] = None,
    query: Optional[str] = None,
    columns: Optional[List[str]] = None,
    limit: int = 50,
    offset: int = 0,
) -> Dict[str, Any]:
    """
    Query, filter, and paginate through Excel rows using Pandas vectorized expressions.
    Example query: "Age > 30 and Status == 'Active'" or "Department in ['Sales', 'Engineering']"
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    limit = max(1, min(limit, 500))
    offset = max(0, offset)

    df = pd.read_excel(file_path, sheet_name=sheet_name or 0, usecols=columns)

    if query:
        try:
            df = df.query(query)
        except Exception as e:
            raise ValueError(f"Invalid query expression '{query}': {str(e)}")

    total_matches = len(df)
    paginated_df = df.iloc[offset : offset + limit]
    clean_df = paginated_df.where(pd.notnull(paginated_df), None)

    return {
        "total_matches": total_matches,
        "offset": offset,
        "limit": limit,
        "returned_rows": len(clean_df),
        "columns": list(df.columns),
        "rows": clean_df.to_dict(orient="records"),
    }


@mcp.tool()
def read_range(
    file_path: str,
    sheet_name: Optional[str] = None,
    start_row: int = 1,
    end_row: int = 50,
    columns: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Read a specific slice of rows (e.g. rows 100 to 200) without loading the entire spreadsheet.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    if start_row < 1:
        start_row = 1
    if end_row < start_row:
        raise ValueError(f"end_row ({end_row}) must be >= start_row ({start_row})")

    nrows = min(end_row - start_row + 1, 1000)

    if start_row <= 1:
        df = pd.read_excel(file_path, sheet_name=sheet_name or 0, nrows=nrows, usecols=columns)
    else:
        header_df = pd.read_excel(file_path, sheet_name=sheet_name or 0, nrows=0, usecols=columns)
        df = pd.read_excel(
            file_path,
            sheet_name=sheet_name or 0,
            skiprows=range(1, start_row),
            nrows=nrows,
            names=header_df.columns,
            usecols=columns,
        )

    clean_df = df.where(pd.notnull(df), None)
    return {
        "start_row": start_row,
        "end_row": start_row + len(clean_df) - 1,
        "row_count": len(clean_df),
        "columns": list(df.columns),
        "rows": clean_df.to_dict(orient="records"),
    }


@mcp.tool()
def summarize_column(
    file_path: str,
    column: str,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Statistical profiling for a specific column in an Excel sheet.
    Computes min, max, median, mean, nulls, or top distinct values.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name or 0, usecols=[column])
    series = df[column]

    total_count = len(series)
    null_count = int(series.isnull().sum())
    unique_count = int(series.nunique())

    summary: Dict[str, Any] = {
        "column": column,
        "total_rows": total_count,
        "null_count": null_count,
        "unique_count": unique_count,
        "data_type": str(series.dtype),
    }

    if pd.api.types.is_numeric_dtype(series):
        valid = series.dropna()
        if len(valid) > 0:
            summary["numeric_stats"] = {
                "min": float(valid.min()),
                "max": float(valid.max()),
                "mean": round(float(valid.mean()), 4),
                "median": float(valid.median()),
                "std": round(float(valid.std()), 4) if len(valid) > 1 else 0.0,
                "sum": round(float(valid.sum()), 4),
            }
    else:
        val_counts = series.value_counts(dropna=True).head(10).to_dict()
        summary["top_values"] = {str(k): int(v) for k, v in val_counts.items()}

    return summary


@mcp.tool()
def search_text(
    file_path: str,
    search_term: str,
    sheet_name: Optional[str] = None,
    max_results: int = 25,
) -> Dict[str, Any]:
    """
    Search across an Excel sheet for a text keyword or number.
    Returns matching row index, column names, and row preview.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name or 0)
    mask = df.astype(str).apply(lambda col: col.str.contains(search_term, case=False, na=False, regex=False))
    matched_rows_mask = mask.any(axis=1)

    matched_indices = df[matched_rows_mask].index[:max_results].tolist()
    results = []

    for idx in matched_indices:
        row_data = df.iloc[idx].where(pd.notnull(df.iloc[idx]), None).to_dict()
        matching_cols = [col for col in df.columns if str(search_term).lower() in str(df.at[idx, col]).lower()]
        results.append({
            "excel_row_number": idx + 2,
            "matching_columns": matching_cols,
            "row_data": row_data,
        })

    return {
        "search_term": search_term,
        "matches_found": int(matched_rows_mask.sum()),
        "returned_matches": len(results),
        "results": results,
    }


@mcp.tool()
def create_workbook(
    file_path: str,
    data: List[Dict[str, Any]],
    sheet_name: str = "Sheet1",
) -> Dict[str, Any]:
    """
    Create a new Excel file from a list of record dictionaries.
    Automatically sets up headers and adjusts column widths.
    """
    if not data:
        raise ValueError("Data list cannot be empty.")

    df = pd.DataFrame(data)
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

    return {
        "status": "success",
        "file_path": file_path,
        "sheet_name": sheet_name,
        "rows_created": len(df),
        "columns": list(df.columns),
    }


@mcp.tool()
def append_rows(
    file_path: str,
    rows: List[Dict[str, Any]],
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Append new rows to an existing sheet without rewriting the entire workbook.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    headers = [cell.value for cell in sheet[1]]
    if not headers or all(h is None for h in headers):
        raise ValueError("Sheet does not have a valid header row to append data to.")

    appended_count = 0
    for row_dict in rows:
        row_values = [row_dict.get(h, None) for h in headers]
        sheet.append(row_values)
        appended_count += 1

    wb.save(file_path)
    wb.close()

    return {
        "status": "success",
        "file_path": file_path,
        "rows_appended": appended_count,
        "new_total_rows": sheet.max_row,
    }


@mcp.tool()
def add_sheet(
    file_path: str,
    sheet_name: str,
    data: Optional[List[Dict[str, Any]]] = None,
    headers: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Add a new worksheet tab to an existing Excel workbook without modifying other sheets.
    Optionally populates it with initial headers and rows.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' already exists in workbook. Existing sheets: {wb.sheetnames}")

    ws = wb.create_sheet(title=sheet_name)
    rows_added = 0

    if data and len(data) > 0:
        cols = headers if headers else list(data[0].keys())
        ws.append(cols)
        for item in data:
            ws.append([item.get(c, None) for c in cols])
            rows_added += 1

        # Auto-adjust column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)
    elif headers:
        ws.append(headers)

    all_sheets = list(wb.sheetnames)
    wb.save(file_path)
    wb.close()

    return {
        "status": "success",
        "file_path": file_path,
        "sheet_added": sheet_name,
        "rows_added": rows_added,
        "all_sheets": all_sheets,
    }


@mcp.tool()
def rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str,
) -> Dict[str, Any]:
    """
    Rename an existing worksheet tab in an Excel workbook.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = openpyxl.load_workbook(file_path)
    if old_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{old_name}' not found. Available sheets: {wb.sheetnames}")
    if new_name in wb.sheetnames:
        raise ValueError(f"A sheet named '{new_name}' already exists.")

    wb[old_name].title = new_name
    all_sheets = list(wb.sheetnames)
    wb.save(file_path)
    wb.close()

    return {
        "status": "success",
        "file_path": file_path,
        "renamed_from": old_name,
        "renamed_to": new_name,
        "all_sheets": all_sheets,
    }


@mcp.tool()
def delete_sheet(
    file_path: str,
    sheet_name: str,
) -> Dict[str, Any]:
    """
    Delete a specific worksheet tab from an existing Excel workbook.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = openpyxl.load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}")
    if len(wb.sheetnames) <= 1:
        raise ValueError("Cannot delete the only sheet in a workbook.")

    wb.remove(wb[sheet_name])
    remaining_sheets = list(wb.sheetnames)
    wb.save(file_path)
    wb.close()

    return {
        "status": "success",
        "file_path": file_path,
        "sheet_deleted": sheet_name,
        "remaining_sheets": remaining_sheets,
    }


@mcp.tool()
def update_cells(
    file_path: str,
    updates: List[Dict[str, Any]],
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Update specific cell coordinates or formulas.
    Example updates: [{"cell": "B5", "value": 1500}, {"cell": "C5", "value": "=A5*B5"}]
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

    updated_count = 0
    for item in updates:
        coord = item.get("cell")
        val = item.get("value")
        if coord:
            sheet[coord] = val
            updated_count += 1

    wb.save(file_path)
    wb.close()

    return {
        "status": "success",
        "file_path": file_path,
        "cells_updated": updated_count,
    }


@mcp.tool()
def export_to_csv(
    file_path: str,
    output_csv_path: str,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Export an Excel sheet to a clean CSV file for fast external processing.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name or 0)
    df.to_csv(output_csv_path, index=False)

    return {
        "status": "success",
        "input_excel": file_path,
        "output_csv": output_csv_path,
        "rows_exported": len(df),
    }


# ==========================================================
# 2. NATIVE MICROSOFT EXCEL WINDOWS AUTOMATION (COM / PyWin32)
# ==========================================================

def _get_abs_path(path: str) -> str:
    return os.path.abspath(os.path.expanduser(path))


@mcp.tool()
def recalculate_and_save(file_path: str) -> Dict[str, Any]:
    """
    [Windows Native Excel] Open workbook in background Microsoft Excel,
    run full formula recalculation (Application.CalculateFull()), save, and close.
    Ensures all dynamic formulas (XLOOKUP, INDEX/MATCH, SUMIFS) are evaluated.
    """
    if not HAS_WIN32:
        raise RuntimeError("PyWin32 is not installed or Windows COM is unavailable.")

    abs_path = _get_abs_path(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found: {abs_path}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_path)
        excel.CalculateFull()
        wb.Save()
        wb.Close()
        return {
            "status": "success",
            "message": "Workbook formulas recalculated and saved using native Microsoft Excel.",
            "file_path": abs_path,
        }
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


@mcp.tool()
def export_to_pdf(
    file_path: str,
    output_pdf_path: str,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    [Windows Native Excel] Export a sheet or entire workbook to a pixel-perfect PDF
    using Microsoft Excel's native print engine.
    """
    if not HAS_WIN32:
        raise RuntimeError("PyWin32 is not installed or Windows COM is unavailable.")

    abs_input = _get_abs_path(file_path)
    abs_output = _get_abs_path(output_pdf_path)

    if not os.path.exists(abs_input):
        raise FileNotFoundError(f"Input file not found: {abs_input}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_input)
        if sheet_name and sheet_name in [s.Name for s in wb.Sheets]:
            ws = wb.Sheets(sheet_name)
            ws.ExportAsFixedFormat(0, abs_output)
        else:
            wb.ExportAsFixedFormat(0, abs_output)

        wb.Close(False)
        return {
            "status": "success",
            "message": f"Successfully exported to PDF: {abs_output}",
            "pdf_path": abs_output,
        }
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


@mcp.tool()
def refresh_data_and_pivots(file_path: str) -> Dict[str, Any]:
    """
    [Windows Native Excel] Refreshes all external data connections, Power Queries,
    and PivotTables in the workbook using native Excel.
    """
    if not HAS_WIN32:
        raise RuntimeError("PyWin32 is not installed or Windows COM is unavailable.")

    abs_path = _get_abs_path(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found: {abs_path}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_path)
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close()
        return {
            "status": "success",
            "message": "All data connections and PivotTables refreshed successfully.",
            "file_path": abs_path,
        }
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


@mcp.tool()
def run_vba_macro(
    file_path: str,
    macro_name: str,
    args: Optional[List[Any]] = None,
) -> Dict[str, Any]:
    """
    [Windows Native Excel] Run a VBA Macro in an Excel workbook (.xlsm or .xlsb).
    """
    if not HAS_WIN32:
        raise RuntimeError("PyWin32 is not installed or Windows COM is unavailable.")

    abs_path = _get_abs_path(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found: {abs_path}")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_path)
        macro_args = args or []
        macro_res = excel.Application.Run(macro_name, *macro_args)
        wb.Save()
        wb.Close()
        return {
            "status": "success",
            "macro_name": macro_name,
            "macro_result": str(macro_res) if macro_res is not None else None,
        }
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


@mcp.tool()
def get_active_excel_window() -> Dict[str, Any]:
    """
    [Windows Native Excel] Check if Microsoft Excel is currently open on the user's
    desktop. If open, returns the active workbook name, active sheet, and selected range.
    """
    if not HAS_WIN32:
        raise RuntimeError("PyWin32 is not installed or Windows COM is unavailable.")

    pythoncom.CoInitialize()
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        ws = excel.ActiveSheet
        sel = excel.Selection

        return {
            "excel_running": True,
            "active_workbook": wb.Name if wb else None,
            "active_workbook_path": wb.FullName if wb else None,
            "active_sheet": ws.Name if ws else None,
            "active_selection": sel.Address if sel else None,
        }
    except Exception:
        return {
            "excel_running": False,
            "message": "No active Excel window currently running on desktop.",
        }
    finally:
        pythoncom.CoUninitialize()


def main():
    mcp.run()


if __name__ == "__main__":
    main()
