import os
import re
import math
import pandas as pd
import numpy as np
from datetime import date, datetime
from google.colab import drive
import dateutil.parser as dparser

# =========================
# CONFIG
# =========================
drive.mount('/content/drive')
your_location=""
INPUT_FOLDER = "/content/drive/MyDrive/"+your_location
OUTPUT_REPORT_XLSX = "/content/drive/MyDrive/cleaned_consistency_report.xlsx"

# How many inconsistent values to show inline in Summary (full list is in the second sheet)
INCONSIST_SAMPLE_SIZE = 10

# =========================
# PATTERN HELPERS
# =========================
# ✅ English: letters + space + underscore + hyphen + dot
RE_ENG = re.compile(r'^[A-Za-z _\-.]+$')

# ✅ Arabic: Arabic letters + space + underscore + hyphen + dot
RE_ARB = re.compile(r'^[\u0600-\u06FF _\-.]+$')

def is_null(x):
    return pd.isna(x) or (isinstance(x, str) and x.strip() == "")

def to_text(x):
    return "" if is_null(x) else str(x)

def is_number(x):
    if is_null(x):
        return False
    if isinstance(x, (int, float, np.integer, np.floating)) and not isinstance(x, bool):
        return not (isinstance(x, float) and (math.isnan(x) or math.isinf(x)))
    try:
        float(str(x))
        return True
    except:
        return False

def is_date(x):
    if is_null(x):
        return False
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return True
    s = str(x).strip()
    has_digit = any(ch.isdigit() for ch in s)
    looks_like_date = any(sep in s for sep in ['-', '/', ':', '.']) or re.search(r'[A-Za-z]{3,}', s)
    if not has_digit or not looks_like_date:
        return False
    try:
        dparser.parse(s, fuzzy=False)
        return True
    except:
        return False

def is_english_words(x):
    if is_null(x):
        return False
    return isinstance(x, str) and bool(RE_ENG.fullmatch(x.strip()))

def is_arabic_words(x):
    if is_null(x):
        return False
    return isinstance(x, str) and bool(RE_ARB.fullmatch(x.strip()))

def detect_pattern_from_first_non_empty(series):
    """
    Return: one of 'number','date','arabic','english','other'
    Based strictly on the FIRST non-empty record.
    """
    for val in series:
        if is_null(val):
            continue
        if is_number(val):
            return 'number'
        if is_date(val):
            return 'date'
        if is_arabic_words(val):
            return 'arabic'
        if is_english_words(val):
            return 'english'
        return 'other'
    return 'other'

def pattern_matches(val, pattern):
    if is_null(val):
        return True
    if pattern == 'number':
        return is_number(val)
    if pattern == 'date':
        return is_date(val)
    if pattern == 'arabic':
        return is_arabic_words(val)
    if pattern == 'english':
        return is_english_words(val)
    return True  # 'other' accepts anything

# =========================
# CONSISTENCY CHECK
# =========================
def evaluate_column(series):
    """
    Returns metrics about pattern + length consistency
    """
    total_records = len(series)
    non_null_mask = ~series.apply(is_null)
    non_null = series[non_null_mask]
    non_null_records = len(non_null)
    null_count = total_records - non_null_records

    if non_null_records == 0:
        return {
            "total_records": int(total_records),
            "non_null_records": 0,
            "null_count": int(null_count),
            "detected_pattern": "other",
            "avg_length": 0,
            "allowed_min_length": 0,
            "allowed_max_length": 0,
            "consistent_count": 0,
            "inconsistent_count": 0,
            "consistency_percentage": 100.0,
            "inconsistent_values": []
        }

    detected_pattern = detect_pattern_from_first_non_empty(series)

    # Compute length stats
    lengths = non_null.apply(lambda x: len(to_text(x)))
    avg_len = lengths.mean()
    min_allowed = avg_len * 0.70   # 30% tolerance
    max_allowed = avg_len * 1.30

    consistent = 0
    inconsistent_values = []

    for idx, val in non_null.items():
        txt_len = len(to_text(val))
        ok_pattern = pattern_matches(val, detected_pattern)

        # ✅ Length check only if detected pattern is "number"
        if detected_pattern == "number":
            ok_length = (min_allowed <= txt_len <= max_allowed)
        else:
            ok_length = True

        if ok_pattern and ok_length:
            consistent += 1
        else:
            inconsistent_values.append(val)

    inconsistent = non_null_records - consistent
    pct = (consistent / non_null_records) * 100 if non_null_records else 100.0

    return {
        "total_records": int(total_records),
        "non_null_records": int(non_null_records),
        "null_count": int(null_count),
        "detected_pattern": detected_pattern,
        "avg_length": round(avg_len, 2),
        "allowed_min_length": math.floor(min_allowed) if not math.isnan(min_allowed) else 0,
        "allowed_max_length": math.ceil(max_allowed) if not math.isnan(max_allowed) else 0,
        "consistent_count": int(consistent),
        "inconsistent_count": int(inconsistent),
        "consistency_percentage": round(pct, 2),
        "inconsistent_values": inconsistent_values
    }

# =========================
# PROCESS FILES
# =========================
def list_excel_files(folder):
    return [os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith('.xlsx') and os.path.isfile(os.path.join(folder, f))]

excel_files = list_excel_files(INPUT_FOLDER)

summary_rows = []
details_rows = []

for file_path in excel_files:
    file_name = os.path.basename(file_path)
    try:
        xfile = pd.ExcelFile(file_path)
        for sheet in xfile.sheet_names:
            df = xfile.parse(sheet_name=sheet, dtype=object)
            if df.empty:
                continue
            for col in df.columns:
                series = df[col]
                metrics = evaluate_column(series)

                # Sample inconsistent values
                sample_list = metrics["inconsistent_values"][:INCONSIST_SAMPLE_SIZE]
                sample_str = " | ".join([str(x) for x in sample_list]) if sample_list else ""

                summary_rows.append({
                    "file": file_name,
                    "sheet": sheet,
                    "column": col,
                    "total_records": metrics["total_records"],
                    "non_null_records": metrics["non_null_records"],
                    "null_count": metrics["null_count"],
                    "detected_pattern": metrics["detected_pattern"],
                    "avg_length": metrics["avg_length"],
                    "allowed_min_length": metrics["allowed_min_length"],
                    "allowed_max_length": metrics["allowed_max_length"],
                    "consistent_count": metrics["consistent_count"],
                    "inconsistent_count": metrics["inconsistent_count"],
                    "consistency_percentage": metrics["consistency_percentage"],
                    "inconsistent_values_sample": sample_str
                })

                if metrics["inconsistent_values"]:
                    for bad in metrics["inconsistent_values"]:
                        details_rows.append({
                            "file": file_name,
                            "sheet": sheet,
                            "column": col,
                            "inconsistent_value": bad
                        })

    except Exception as e:
        summary_rows.append({
            "file": file_name,
            "sheet": "<error>",
            "column": "<error>",
            "total_records": "",
            "non_null_records": "",
            "null_count": "",
            "detected_pattern": "<error>",
            "avg_length": "",
            "allowed_min_length": "",
            "allowed_max_length": "",
            "consistent_count": "",
            "inconsistent_count": "",
            "consistency_percentage": "",
            "inconsistent_values_sample": f"Error: {e}"
        })
        details_rows.append({
            "file": file_name,
            "sheet": "<error>",
            "column": "<error>",
            "inconsistent_value": f"Error: {e}"
        })

# =========================
# SAVE REPORT
# =========================
summary_df = pd.DataFrame(summary_rows)
details_df = pd.DataFrame(details_rows)

summary_order = [
    "file", "sheet", "column",
    "total_records", "non_null_records", "null_count",
    "detected_pattern",
    "avg_length", "allowed_min_length", "allowed_max_length",
    "consistent_count", "inconsistent_count", "consistency_percentage",
    "inconsistent_values_sample"
]
summary_df = summary_df.reindex(columns=summary_order)

with pd.ExcelWriter(OUTPUT_REPORT_XLSX, engine="openpyxl") as writer:
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    details_df.to_excel(writer, sheet_name="Inconsistent Values", index=False)

print(f"✅ Report saved: {OUTPUT_REPORT_XLSX}")
