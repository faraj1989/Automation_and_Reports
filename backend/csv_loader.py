#!/usr/bin/env python3
"""
Libyana NPM - CSV Loader
Universal CSV reader with auto-detection of header row and delimiter.
Works for all CSV types: 2G, 3G, 4G, User KPIs, Traffic, etc.
"""

import os
import logging
import pandas as pd

logger = logging.getLogger(__name__)

# Huawei PM export null markers. "/0" means the KPI's ratio had zero
# attempts in that hour (counter divide-by-zero) - it is NOT the same as a
# 0% success rate, which means attempts were made and all failed. "NIL"
# means the counter was not collected/supported for that cell. Both must
# become NaN, never 0 - treating them as 0 would flag a quiet, healthy cell
# as a critical failure.
NULL_TOKENS = ['/0', 'NIL', 'NULL', 'N/A', '--']


def clean_null_tokens(df):
    """Replace Huawei's non-numeric null markers with NaN, then coerce any
    column that is now fully numeric back to a numeric dtype (read_csv left
    it as object/string purely because of the stray markers)."""
    df = df.replace(NULL_TOKENS, pd.NA)
    for col in df.columns:
        # Newer pandas may read text columns as object OR as a string dtype
        # (e.g. "string[pyarrow]") depending on backend config - checking
        # "not already numeric" catches both, instead of assuming object.
        if pd.api.types.is_numeric_dtype(df[col]):
            continue
        coerced = pd.to_numeric(df[col], errors='coerce')
        # Only swap in the numeric version if this didn't just turn
        # genuine text (site names, "Whole Network", "100%", ...) into
        # NaN - that would mean the column was never numeric to begin with.
        still_has_text = df[col].notna() & coerced.isna()
        if not still_has_text.any():
            df[col] = coerced
    return df


def detect_delimiter(filepath):
    """
    Detect the delimiter used in the CSV file.
    Tries: comma, tab, semicolon, pipe.
    Returns the most likely delimiter.
    """
    delimiters = [',', '\t', ';', '|']
    best_delimiter = ','
    best_count = 0

    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            # Read first 10 lines
            lines = []
            for i, line in enumerate(f):
                if i >= 10:
                    break
                if line.strip():
                    lines.append(line.strip())

            if not lines:
                return ','

            for delim in delimiters:
                count = 0
                for line in lines:
                    count += line.count(delim)
                if count > best_count:
                    best_count = count
                    best_delimiter = delim

    except Exception as e:
        logger.warning(f"Could not detect delimiter: {e}")

    return best_delimiter


def find_header_row(filepath, header_keywords=None):
    """
    Find the header row by scanning the file.
    Returns the row index where the header starts.
    """
    if header_keywords is None:
        header_keywords = ['Time', 'Date']

    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for i, line in enumerate(f):
                stripped = line.strip()
                # Check if the line starts with any of the header keywords
                for keyword in header_keywords:
                    if stripped.startswith(keyword):
                        return i
        return None
    except Exception as e:
        logger.error(f"Failed to scan file {filepath}: {e}")
        return None


def read_csv_skip_metadata(filepath, header_keywords=None, delimiter=None):
    """
    Universal CSV reader that auto-detects header row and delimiter.

    Args:
        filepath: Path to the CSV file
        header_keywords: List of keywords to identify header row (default: ['Time', 'Date'])
        delimiter: Force a specific delimiter (auto-detect if None)

    Returns:
        DataFrame or None if reading fails
    """
    if not os.path.exists(filepath):
        logger.error(f"File not found: {filepath}")
        return None

    # Detect delimiter if not provided
    if delimiter is None:
        delimiter = detect_delimiter(filepath)
        logger.debug(f"Detected delimiter: '{repr(delimiter)}' for {os.path.basename(filepath)}")

    # Find header row
    if header_keywords is None:
        header_keywords = ['Time', 'Date']

    skip_rows = find_header_row(filepath, header_keywords)

    if skip_rows is None:
        logger.error(f"Could not find header row in {os.path.basename(filepath)}")
        # Try to read without skipping (as fallback)
        skip_rows = 0

    logger.info(f"Header found at row {skip_rows} in {os.path.basename(filepath)}")

    # Read the CSV
    try:
        df = pd.read_csv(
            filepath,
            skiprows=skip_rows,
            encoding='utf-8',
            delimiter=delimiter,
            on_bad_lines='skip'
        )
    except Exception as e:
        # Try with different encoding
        try:
            df = pd.read_csv(
                filepath,
                skiprows=skip_rows,
                encoding='latin-1',
                delimiter=delimiter,
                on_bad_lines='skip'
            )
        except Exception as e2:
            logger.error(f"Failed to read CSV {filepath}: {e2}")
            return None

    if df.empty:
        logger.warning(f"Empty DataFrame after reading {filepath}")
        return df

    # Clean column names (strip spaces)
    df.columns = df.columns.str.strip()

    # Normalize Huawei's "/0" and "NIL" null markers to NaN across all columns
    null_counts = {col: int((df[col].astype(str).isin(NULL_TOKENS)).sum()) for col in df.columns}
    total_nulls = sum(null_counts.values())
    if total_nulls > 0:
        df = clean_null_tokens(df)
        logger.info(f"🧹 Normalized {total_nulls} '/0' or 'NIL' markers to NaN in {os.path.basename(filepath)}")

    logger.info(f"Loaded {len(df)} rows, {len(df.columns)} columns from {os.path.basename(filepath)}")

    # Drop footer row (if the last row's first column is not a valid date/time)
    try:
        last_value = str(df.iloc[-1, 0]).strip()
        pd.to_datetime(last_value)
        # If it parses, it's data, not footer
    except (ValueError, TypeError, IndexError):
        if len(df) > 0:
            df = df.iloc[:-1]
            logger.debug(f"Dropped footer row from {filepath}")

    return df


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        print(f"Testing: {test_file}")
        df = read_csv_skip_metadata(test_file)
        if df is not None:
            print(f"✅ Loaded {len(df)} rows, {len(df.columns)} columns")
            print(f"Columns: {df.columns.tolist()}")
            print("\nFirst 3 rows:")
            print(df.head(3))
        else:
            print("❌ Failed to load")