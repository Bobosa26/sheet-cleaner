#!/usr/bin/env python3
"""Sheet Cleaner: merge, clean and summarise messy CSV/Excel files.

Usage:
    python clean_sheets.py sample_data -o report.xlsx
    python clean_sheets.py my_folder -o report.xlsx --group-by region --sum amount
"""
import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

SUPPORTED = {".csv", ".xlsx", ".xls"}
PHONE_HINTS = ("phone", "tel", "mobile", "contact")
AMOUNT_HINTS = ("amount", "price", "total", "cost", "revenue", "salary", "paid")
DATE_HINTS = ("date",)
TITLE_HINTS = ("name", "customer", "product", "category", "region", "town", "district")


# ---------- small cleaning helpers ----------

def snake(name):
    """'Customer Name ' -> 'customer_name'."""
    text = re.sub(r"[^0-9a-zA-Z]+", "_", str(name).strip().lower()).strip("_")
    return text or "column"


def clean_phone(value):
    """Normalise Ugandan numbers to +256XXXXXXXXX. Returns None if invalid."""
    if value is None or pd.isna(value):
        return None
    digits = re.sub(r"\D", "", str(value))
    if digits.startswith("256") and len(digits) == 12:
        pass
    elif digits.startswith("0") and len(digits) == 10:
        digits = "256" + digits[1:]
    elif len(digits) == 9:
        digits = "256" + digits
    else:
        return None
    return "+" + digits


def clean_amount(value):
    """'UGX 1,200,000' -> 1200000.0. Returns None if it can't be parsed."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^\d.\-]", "", str(value).replace(",", ""))
    try:
        return float(text)
    except ValueError:
        return None


def clean_date(value):
    """Parse dates; ISO (2026-09-05) is read year-first, others day-first (03/09/2026 = 3 Sept)."""
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value
    text = str(value).strip()
    iso = re.match(r"^\d{4}-\d{1,2}-\d{1,2}", text)
    return pd.to_datetime(text, dayfirst=not iso, errors="coerce")


def clean_text(value):
    """Trim, collapse spaces, turn empty strings into None."""
    if value is None or pd.isna(value):
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def has_hint(column, hints):
    return any(h in column for h in hints)


# ---------- reading ----------

def read_file(path):
    """Read a CSV or every sheet of an Excel file into one DataFrame."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(path, dtype=object)
        except UnicodeDecodeError:
            df = pd.read_csv(path, dtype=object, encoding="latin-1")
    else:
        sheets = pd.read_excel(path, sheet_name=None, dtype=object)
        df = pd.concat(sheets.values(), ignore_index=True)
    df["source_file"] = path.name
    return df


def find_files(folder):
    folder = Path(folder)
    if folder.is_file():
        return [folder]
    return sorted(p for p in folder.iterdir() if p.suffix.lower() in SUPPORTED)


# ---------- cleaning one table ----------

def clean_frame(df, issues, source):
    """Clean one DataFrame. Problems found are appended to `issues`."""
    df = df.copy()

    # 1. tidy column names (and make them unique)
    names, seen = [], {}
    for col in df.columns:
        name = snake(col)
        seen[name] = seen.get(name, 0) + 1
        names.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    df.columns = names

    # 2. trim text everywhere, blanks become missing
    for col in df.columns:
        df[col] = df[col].map(clean_text)

    # 3. drop rows that are completely empty
    data_cols = [c for c in df.columns if c != "source_file"]
    empty = df[data_cols].isna().all(axis=1)
    if empty.any():
        issues.append((source, "empty rows removed", int(empty.sum())))
        df = df[~empty]

    # 4. column-specific cleaning
    for col in data_cols:
        before_missing = df[col].isna().sum()
        if has_hint(col, PHONE_HINTS):
            df[col] = df[col].map(clean_phone)
            bad = df[col].isna().sum() - before_missing
            if bad:
                issues.append((source, f"invalid phone numbers in '{col}'", int(bad)))
        elif has_hint(col, AMOUNT_HINTS):
            df[col] = df[col].map(clean_amount)
            bad = df[col].isna().sum() - before_missing
            if bad:
                issues.append((source, f"unreadable amounts in '{col}'", int(bad)))
        elif has_hint(col, DATE_HINTS):
            parsed = pd.to_datetime(df[col].map(clean_date), errors="coerce")
            bad = parsed.isna().sum() - before_missing
            if bad:
                issues.append((source, f"unreadable dates in '{col}'", int(bad)))
            df[col] = parsed
        elif has_hint(col, TITLE_HINTS):
            df[col] = df[col].map(lambda v: v.title() if isinstance(v, str) else v)

    return df.reset_index(drop=True)


# ---------- merging and summarising ----------

def merge_frames(frames, issues):
    merged = pd.concat(frames, ignore_index=True)
    key_cols = [c for c in merged.columns if c != "source_file"]
    dupes = merged.duplicated(subset=key_cols, keep="first")
    if dupes.any():
        issues.append(("all files", "duplicate rows removed", int(dupes.sum())))
        merged = merged[~dupes]
    return merged.reset_index(drop=True)


def guess_column(df, hints):
    for col in df.columns:
        if has_hint(col, hints):
            return col
    return None


def summarise(df, group_by=None, sum_col=None):
    """Total of `sum_col` for each value of `group_by` (both auto-detected if omitted)."""
    group_by = group_by or guess_column(df, ("category", "product", "region", "town", "district"))
    sum_col = sum_col or guess_column(df, AMOUNT_HINTS)
    if group_by not in df.columns or sum_col not in df.columns:
        return pd.DataFrame({"note": ["No group/amount column found. Use --group-by and --sum."]})
    out = (
        df.groupby(group_by, dropna=False)[sum_col]
        .agg(rows="count", total="sum")
        .reset_index()
        .sort_values("total", ascending=False)
    )
    return out


# ---------- writing the report ----------

def write_report(df, summary, issues, out_path):
    issues_df = pd.DataFrame(issues, columns=["file", "problem", "count"])
    if issues_df.empty:
        issues_df = pd.DataFrame({"file": ["-"], "problem": ["No problems found"], "count": [0]})

    with pd.ExcelWriter(out_path, engine="openpyxl", datetime_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        issues_df.to_excel(writer, sheet_name="Issues Log", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for column in ws.columns:
                longest = max(len(str(c.value)) if c.value is not None else 0 for c in column)
                ws.column_dimensions[column[0].column_letter].width = min(longest + 2, 40)


def run(input_path, out_path, group_by=None, sum_col=None):
    files = find_files(input_path)
    if not files:
        raise SystemExit(f"No CSV/Excel files found in {input_path}")
    issues, frames = [], []
    for f in files:
        frames.append(clean_frame(read_file(f), issues, f.name))
    merged = merge_frames(frames, issues)
    summary = summarise(merged, group_by, sum_col)
    write_report(merged, summary, issues, out_path)
    return merged, summary, issues


def main(argv=None):
    p = argparse.ArgumentParser(description="Clean, merge and summarise CSV/Excel files.")
    p.add_argument("input", help="folder (or single file) with CSV/Excel files")
    p.add_argument("-o", "--output", default="report.xlsx", help="output Excel file")
    p.add_argument("--group-by", help="column to group the summary by")
    p.add_argument("--sum", dest="sum_col", help="numeric column to total")
    args = p.parse_args(argv)

    merged, summary, issues = run(args.input, args.output, args.group_by, args.sum_col)
    print(f"Done: {len(merged)} clean rows, {len(issues)} issue types logged -> {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
