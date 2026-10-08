import argparse
import hashlib
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import load_env_file
from project_logging import setup_logger as create_logger

PROJECT_ROOT = Path(__file__).resolve().parent
load_env_file()

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}
RECOVERABLE_EXTENSIONS = SUPPORTED_EXTENSIONS | {".zip"}
COMPREHENSIVE_PREFIX = "Comprehensive_Analysis"
PROCESSED_LEDGER_FILENAME = "processed_files_ledger.json"
EXPORT_ROW_CAP = 100_000     # SmartCare's hard limit per Excel(All Data) export
MIN_DAY_ROWS_SHARE = 0.9     # never replace a stored day with one that has <90% of its rows

# =============================================================
# CENTRALIZED DIRECTORIES (from GUI)
# =============================================================
DATA_ROOT = Path(os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\NOC_Data"))

# SmartCare paths
SMARTCARE_DOWNLOAD_DIR = Path(os.environ.get("SMARTCARE_DOWNLOAD_DIR", DATA_ROOT / "Downloads"))
SMARTCARE_OUTPUT_DIR = Path(os.environ.get("SMARTCARE_OUTPUT_DIR", DATA_ROOT / "SmartCare_Exports"))

# Analysis paths
ANALYSIS_SOURCE_DIR = Path(os.environ.get("ANALYSIS_SOURCE_DIR", SMARTCARE_OUTPUT_DIR))
ANALYSIS_OUTPUT_DIR = Path(os.environ.get("ANALYSIS_OUTPUT_DIR", DATA_ROOT / "Processed_Analysis"))
ANALYSIS_HISTORY_FILE = Path(
    os.environ.get("ANALYSIS_HISTORY_FILE", ANALYSIS_OUTPUT_DIR / "Comprehensive_Analysis_Historical.xlsx"))


# =============================================================


def setup_logger() -> logging.Logger:
    return create_logger("download-analysis")


def normalize_column_name(name: object) -> str:
    if pd.isna(name):
        return "column"
    text = str(name).strip()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-z0-9_]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return (text or "column").lower()


def load_dataframe(file_path: Path) -> pd.DataFrame:
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(
            file_path,
            encoding="utf-8-sig",
            engine="python",
            na_values=["", "NA", "N/A", "null", "None"],
            keep_default_na=True,
        )
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(file_path, engine="openpyxl")
    raise ValueError(f"Unsupported file type: {file_path}")


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [normalize_column_name(col) for col in cleaned.columns]

    for col in cleaned.columns:
        cleaned[col] = cleaned[col].replace({"": pd.NA})

    return cleaned


def parse_time_column(df: pd.DataFrame) -> pd.Series:
    if "time" not in df.columns:
        raise ValueError("Expected a Time column in the input data.")

    series = pd.to_datetime(df["time"], errors="coerce", utc=False)
    if series.isna().all():
        return pd.Series(pd.NA, index=df.index)

    return series.dt.date.astype(str)


def normalize_rate_series(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.strip().replace("--", pd.NA)
    cleaned = cleaned.str.replace("%", "", regex=False)
    return pd.to_numeric(cleaned, errors="coerce")


def build_top100_traffic(df: pd.DataFrame) -> pd.DataFrame:
    if "application" not in df.columns or "total_traffic_byte" not in df.columns:
        raise ValueError("Expected Application and Total Traffic(Byte) columns in the input data.")

    df = df.copy()
    df["date"] = parse_time_column(df)
    df["total_traffic_byte"] = pd.to_numeric(df["total_traffic_byte"], errors="coerce").fillna(0)
    grouped = (
        df.groupby(["date", "application"], dropna=False, as_index=False)["total_traffic_byte"].sum()
        .rename(columns={"total_traffic_byte": "total_traffic_bytes"})
    )
    grouped["total_traffic_gb"] = grouped["total_traffic_bytes"] / 1024 ** 3
    grouped = grouped.sort_values(["date", "total_traffic_bytes"], ascending=[True, False])
    top100 = grouped.groupby("date", group_keys=False).head(100)
    return top100


def build_top10_weekly_applications(top100_df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Top-N applications per week, rolled up from the daily Top100 sheet
    (per-week sum of each application's daily totals) rather than re-reading
    raw per-record data - an intentional approximation: an application that
    never cracked a given day's top 100 doesn't contribute to that day's
    total here, same as it wouldn't show up in the Top100 sheet itself.
    Week = the Monday that starts each date's ISO week, so weeks sort and
    group the same way regardless of which day of the week a report runs."""
    columns = ["week", "week_number", "days_in_week", "application", "total_traffic_bytes", "total_traffic_gb"]
    if top100_df.empty:
        return pd.DataFrame(columns=columns)
    working = top100_df.copy()
    parsed_date = pd.to_datetime(working["date"], errors="coerce")
    working["week"] = (parsed_date - pd.to_timedelta(parsed_date.dt.weekday, unit="D")).dt.strftime("%Y-%m-%d")
    weekly = working.groupby(["week", "application"], dropna=False, as_index=False)["total_traffic_bytes"].sum()
    # How many days of data each week's totals are built from - a week with
    # missing days (export gaps) has smaller totals and must not be read as a
    # real week-over-week drop.
    weekly["days_in_week"] = weekly["week"].map(working.groupby("week")["date"].nunique()).astype(int)
    weekly["total_traffic_gb"] = weekly["total_traffic_bytes"] / 1024 ** 3
    # ISO week number of the week-start date - repeats every year, but this
    # archive only spans one, so a bare number (not "YYYY-Www") stays readable.
    weekly["week_number"] = pd.to_datetime(weekly["week"]).dt.isocalendar().week.astype(int)
    weekly = weekly.sort_values(["week", "total_traffic_bytes"], ascending=[False, False])
    top10 = weekly.groupby("week", group_keys=False).head(top_n)
    return top10[columns]


def build_rate_metrics(df: pd.DataFrame) -> pd.DataFrame:
    rate_columns = [
        "tcp_connection_success_rate",
        "downlink_tcp_retransmission_rate",
        "average_tcp_packet_loss_rate",
        "downlink_tcp_packet_loss_rate",
        "tcp_connection_success_rate_included_rst",
    ]

    missing = [col for col in ["time"] + rate_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missing expected rate columns: {missing}")

    rates = df[["time"] + rate_columns + ["total_traffic_byte"]].copy()
    rates["date"] = parse_time_column(rates)
    for col in rate_columns:
        rates[col] = normalize_rate_series(rates[col])

    rates["total_traffic_bytes"] = pd.to_numeric(rates["total_traffic_byte"], errors="coerce").fillna(0)
    rates["total_traffic_gb"] = rates["total_traffic_bytes"] / 1024 ** 3

    daily = (
        rates.groupby("date", dropna=False, as_index=False)
        .agg(
            tcp_connection_success_rate=("tcp_connection_success_rate", "mean"),
            downlink_tcp_retransmission_rate=("downlink_tcp_retransmission_rate", "mean"),
            average_tcp_packet_loss_rate=("average_tcp_packet_loss_rate", "mean"),
            downlink_tcp_packet_loss_rate=("downlink_tcp_packet_loss_rate", "mean"),
            tcp_connection_success_rate_included_rst=("tcp_connection_success_rate_included_rst", "mean"),
            total_traffic_gb=("total_traffic_gb", "sum"),
            source_rows=("date", "size"),
        )
    )

    # The rate columns above are a plain mean across ~3,800 application rows
    # per day, so thousands of tiny apps outweigh Facebook/TikTok/YouTube and
    # e.g. DL retransmission reads ~3x worse than users actually see (2.8% vs
    # 0.8%, 2026-09-28..10-04). Kept unchanged for trend continuity (days
    # before 2026-08-06 have no raw export left to recompute); these
    # *_weighted columns are the network-true figures. Packet-based rates are
    # weighted by downlink packets (rate = lost/retransmitted packets / packets,
    # so this is exact for the DL rates and the closest available proxy for the
    # UL+DL average); connection success rates by traffic, since connection
    # counts aren't in the export.
    packets = pd.to_numeric(df.get("downlink_packets_packets"), errors="coerce") \
        if "downlink_packets_packets" in df.columns else None
    weights = {
        "tcp_connection_success_rate": rates["total_traffic_bytes"],
        "tcp_connection_success_rate_included_rst": rates["total_traffic_bytes"],
        "downlink_tcp_retransmission_rate": packets,
        "average_tcp_packet_loss_rate": packets,
        "downlink_tcp_packet_loss_rate": packets,
    }
    for col, w in weights.items():
        if w is None:
            continue
        w = w.reindex(rates.index).fillna(0)
        ok = rates[col].notna() & (w > 0)
        num = (rates[col] * w).where(ok, 0).groupby(rates["date"]).sum()
        den = w.where(ok, 0).groupby(rates["date"]).sum()
        daily[f"{col}_weighted"] = daily["date"].map(num / den.replace(0, pd.NA))

    return daily


def build_summary(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    rows = []
    for col in df.columns:
        series = df[col]
        non_null = int(series.notna().sum())
        missing = int(series.isna().sum())
        dtype = str(series.dtype)
        sample = ""
        if not series.dropna().empty:
            sample = str(series.dropna().iloc[0])
        rows.append(
            {
                "source_file": source_name,
                "column_name": col,
                "dtype": dtype,
                "rows": int(len(df)),
                "non_null_values": non_null,
                "missing_values": missing,
                "sample_value": sample,
            }
        )

    return pd.DataFrame(rows)


def file_fingerprint(path: Path) -> str:
    """Content hash so a re-downloaded or relocated copy of a file already folded
    into the historical archive is recognized even if its name or mtime changed."""
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_processed_ledger(history_path: Path) -> dict:
    ledger_path = history_path.parent / PROCESSED_LEDGER_FILENAME
    if not ledger_path.exists():
        return {}
    try:
        return json.loads(ledger_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def save_processed_ledger(history_path: Path, ledger: dict) -> None:
    ledger_path = history_path.parent / PROCESSED_LEDGER_FILENAME
    ledger_path.parent.mkdir(parents=True, exist_ok=True)
    ledger_path.write_text(json.dumps(ledger, indent=2), encoding="utf-8")


def recover_stray_downloads(search_dirs: List[Path], target_dir: Path, logger: logging.Logger) -> List[Path]:
    """Sweep up Comprehensive_Analysis exports that never made it into target_dir:
    files left behind in a raw browser-download folder, or a file sitting one level
    deep inside its own subfolder of target_dir (e.g. an old manual extraction).
    Duplicates (identical content already present at the destination) are removed
    instead of copied again."""
    target_dir = Path(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    target_resolved = target_dir.resolve()
    recovered: List[Path] = []
    seen_dirs = set()

    for search_dir in search_dirs:
        search_dir = Path(search_dir)
        if not search_dir.exists():
            continue
        search_resolved = search_dir.resolve()
        if search_resolved in seen_dirs:
            continue
        seen_dirs.add(search_resolved)

        for entry in list(search_dir.iterdir()):
            if entry.is_file():
                candidates = [entry]
            elif entry.is_dir() and entry.resolve() != target_resolved:
                candidates = [p for p in entry.iterdir() if p.is_file()]
            else:
                continue

            for candidate in candidates:
                if not candidate.stem.startswith(COMPREHENSIVE_PREFIX):
                    continue
                if candidate.suffix.lower() not in RECOVERABLE_EXTENSIONS:
                    continue
                if candidate.resolve().parent == target_resolved:
                    continue  # already sitting flat where it belongs

                destination = target_dir / candidate.name
                if destination.exists():
                    if file_fingerprint(destination) == file_fingerprint(candidate):
                        logger.info(f"Duplicate of an already-recovered export, discarding stray copy: {candidate}")
                        candidate.unlink()
                        continue
                    destination = destination.with_name(
                        f"{destination.stem}_{datetime.now():%Y%m%d_%H%M%S}{destination.suffix}"
                    )

                logger.info(f"Recovering old/unprocessed export: {candidate} -> {destination}")
                shutil.move(str(candidate), str(destination))
                recovered.append(destination)

            if entry.is_dir() and entry.resolve() != target_resolved:
                try:
                    if not any(entry.iterdir()):
                        entry.rmdir()
                except OSError:
                    pass

    return recovered


def load_existing_dataframe(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig")
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, engine="openpyxl")
    return pd.DataFrame()


def load_existing_excel_sheet(excel_path: Path, sheet_name: str, legacy_name: str = None) -> pd.DataFrame:
    """legacy_name lets a renamed sheet (e.g. "Top100" -> "Top100 per day")
    pick up its history from the old sheet name once, the first time this
    runs after the rename, instead of silently starting over empty - the
    write path always uses the new name, so the old name naturally drops out
    of the workbook after that one run."""
    if not excel_path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        pass
    if legacy_name:
        try:
            return pd.read_excel(excel_path, sheet_name=legacy_name, engine="openpyxl")
        except Exception:
            pass
    return pd.DataFrame()


def merge_records(existing: pd.DataFrame, new_data: pd.DataFrame, subset: List[str]) -> pd.DataFrame:
    if existing.empty:
        return new_data
    combined = pd.concat([existing, new_data], ignore_index=True)
    combined = combined.drop_duplicates(subset=subset, keep="last")
    return combined


def save_outputs(df: pd.DataFrame, summary_df: pd.DataFrame, output_dir: Path, base_name: str) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_path = output_dir / f"{base_name}_analysis.xlsx"
    csv_path = output_dir / f"{base_name}_data.csv"
    summary_path = output_dir / f"{base_name}_summary.csv"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Processed_Data", index=False)
        summary_df.to_excel(writer, sheet_name="Column_Summary", index=False)

    df.to_csv(csv_path, index=False)
    summary_df.to_csv(summary_path, index=False)
    return [excel_path, csv_path, summary_path]


def save_historical_archive(history_path: Path, top100_df: pd.DataFrame, metrics_df: pd.DataFrame) -> Path:
    history_path.parent.mkdir(parents=True, exist_ok=True)
    existing_top100 = load_existing_excel_sheet(history_path, "Top100 per day", legacy_name="Top100")
    existing_metrics = load_existing_excel_sheet(history_path, "Metrics")

    # Never replace a stored day with a clearly smaller (partial) version of it.
    # Only comparable when the stored day recorded its source_rows (days
    # archived before 2026-10-06 didn't).
    if (not existing_metrics.empty and "source_rows" in existing_metrics.columns
            and "source_rows" in metrics_df.columns):
        stored = existing_metrics.assign(date=existing_metrics["date"].astype(str)) \
            .dropna(subset=["source_rows"]).set_index("date")["source_rows"]
        new_rows = metrics_df.assign(date=metrics_df["date"].astype(str)).set_index("date")["source_rows"]
        keep_stored = [d for d, n in new_rows.items()
                       if d in stored.index and n < MIN_DAY_ROWS_SHARE * stored[d]]
        if keep_stored:
            logging.getLogger("download-analysis").warning(
                f"Kept the stored version of {keep_stored}: this export has <{MIN_DAY_ROWS_SHARE:.0%} "
                f"of the rows for those day(s), i.e. they look partial")
            top100_df = top100_df[~top100_df["date"].astype(str).isin(keep_stored)]
            metrics_df = metrics_df[~metrics_df["date"].astype(str).isin(keep_stored)]

    date_values = set(top100_df["date"].astype(str).unique())
    if not existing_top100.empty and "date" in existing_top100.columns:
        existing_top100 = existing_top100[~existing_top100["date"].astype(str).isin(date_values)]
    if not existing_metrics.empty and "date" in existing_metrics.columns:
        existing_metrics = existing_metrics[~existing_metrics["date"].astype(str).isin(date_values)]

    merged_top100 = pd.concat([existing_top100, top100_df],
                              ignore_index=True) if not existing_top100.empty else top100_df.copy()
    merged_metrics = pd.concat([existing_metrics, metrics_df],
                               ignore_index=True) if not existing_metrics.empty else metrics_df.copy()

    if not merged_top100.empty:
        merged_top100 = merged_top100.sort_values(["date", "total_traffic_bytes"], ascending=[False, False])
    # Recomputed from the FULL merged history every run (not merged
    # incrementally) so a week whose daily data changes always reflects
    # correct totals, never a stale partial-week rollup.
    weekly_top10 = build_top10_weekly_applications(merged_top100)
    if not merged_metrics.empty:
        merged_metrics = merged_metrics.sort_values("date", ascending=False)

    with pd.ExcelWriter(history_path, engine="openpyxl") as writer:
        if not merged_top100.empty:
            merged_top100.to_excel(writer, sheet_name="Top100 per day", index=False)
        if not weekly_top10.empty:
            weekly_top10.to_excel(writer, sheet_name="Top10 Application per week", index=False)
        if not merged_metrics.empty:
            merged_metrics.to_excel(writer, sheet_name="Metrics", index=False)

    return history_path


def save_comprehensive_outputs(
        top100_df: pd.DataFrame,
        metrics_df: pd.DataFrame,
        output_dir: Path,
        base_name: str,
        history_path: Path,
) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_path = output_dir / f"{base_name}_report.xlsx"
    top_csv = output_dir / f"{base_name}_top100.csv"
    metrics_csv = output_dir / f"{base_name}_metrics.csv"
    history_path = Path(history_path)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        top100_df.to_excel(writer, sheet_name="Top100_Traffic", index=False)
        metrics_df.to_excel(writer, sheet_name="Rate_Metrics", index=False)

    top100_df.to_csv(top_csv, index=False)
    metrics_df.to_csv(metrics_csv, index=False)

    historical_file = save_historical_archive(history_path, top100_df, metrics_df)
    return [excel_path, top_csv, metrics_csv, historical_file]


def process_data_file(file_path: Path, output_dir: Path, history_path: Path, logger: logging.Logger,
                       ledger: dict) -> dict:
    if not file_path.stem.startswith(COMPREHENSIVE_PREFIX):
        logger.info(f"Skipping file because it does not start with {COMPREHENSIVE_PREFIX}: {file_path.name}")
        return {}

    fingerprint = file_fingerprint(file_path)
    already = ledger.get(fingerprint)
    if already:
        logger.info(
            f"Skipping {file_path.name}: identical content already in the historical archive "
            f"(processed {already.get('processed_at')} as {already.get('source_name')})"
        )
        return {}

    logger.info(f"Processing {file_path}")
    df = load_dataframe(file_path)
    cleaned_df = clean_dataframe(df)

    # SmartCare cuts every export at EXPORT_ROW_CAP rows (~26 days at ~3,850
    # rows/day). An export that hit the cap stops mid-day, so its last day is
    # partial - drop it rather than let it replace a complete stored day
    # (what happened to 2026-09-01 on 2026-10-06).
    if len(cleaned_df) >= EXPORT_ROW_CAP:
        day = parse_time_column(cleaned_df)
        last_day = day.dropna().max()
        cleaned_df = cleaned_df[day != last_day]
        logger.warning(f"{file_path.name} hit the {EXPORT_ROW_CAP:,}-row export cap - dropped its last, "
                       f"partial day {last_day}. Export long ranges in chunks of <= 25 days.")
    top100_df = build_top100_traffic(cleaned_df)
    metrics_df = build_rate_metrics(cleaned_df)

    base_name = re.sub(r"[^A-Za-z0-9._-]+", "_", file_path.stem).strip("_")
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs = save_comprehensive_outputs(top100_df, metrics_df, output_dir, base_name, history_path)

    ledger[fingerprint] = {
        "source_name": file_path.name,
        "processed_at": datetime.now().isoformat(timespec="seconds"),
        "rows": int(len(cleaned_df)),
    }

    return {
        "source": str(file_path),
        "rows": int(len(cleaned_df)),
        "columns": int(len(cleaned_df.columns)),
        "output_dir": str(output_dir),
        "outputs": [str(p) for p in outputs],
    }


def safe_rmtree(path: Path, logger: logging.Logger) -> None:
    def onerror(func, path_str, exc_info):
        import stat

        exc_type, exc_value, _ = exc_info
        if exc_type is PermissionError:
            logger.warning(f"Permission error removing {path_str}; retrying with readonly flag")
            os.chmod(path_str, stat.S_IWRITE)
            func(path_str)
        else:
            raise

    shutil.rmtree(path, onerror=onerror)


def safe_extract_zip(archive: zipfile.ZipFile, destination: Path) -> None:
    """Extract an archive only when every member stays within destination."""
    destination = destination.resolve()
    for member in archive.infolist():
        target = (destination / member.filename).resolve()
        if target != destination and destination not in target.parents:
            raise ValueError(f"Unsafe ZIP member path: {member.filename}")
    archive.extractall(destination)


def process_zip_file(zip_path: Path, output_root: Path, history_path: Path, logger: logging.Logger,
                     ledger: dict) -> List[dict]:
    if not zip_path.stem.startswith(COMPREHENSIVE_PREFIX):
        logger.info(f"Skipping archive because it does not start with {COMPREHENSIVE_PREFIX}: {zip_path.name}")
        return []

    extract_dir = output_root / "extracted" / zip_path.stem
    if extract_dir.exists():
        try:
            safe_rmtree(extract_dir, logger)
        except Exception as exc:
            logger.warning(f"Could not remove existing extract directory: {exc}")
    extract_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"Extracting archive: {zip_path}")
    with zipfile.ZipFile(zip_path, "r") as archive:
        safe_extract_zip(archive, extract_dir)

    data_files = sorted(
        [p for p in extract_dir.rglob("*") if
         p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS and p.stem.startswith(COMPREHENSIVE_PREFIX)],
        key=lambda p: p.stat().st_size,
        reverse=True,
    )

    if not data_files:
        logger.warning(f"No supported Comprehensive_Analysis files found inside {zip_path}")
        return []

    results = []
    for data_file in data_files:
        per_file_output_dir = output_root / "processed" / data_file.stem
        result = process_data_file(data_file, per_file_output_dir, history_path, logger, ledger)
        if result:
            results.append(result)
    return results


def discover_files(input_dir: Path) -> List[Path]:
    if not input_dir.exists():
        return []

    # Look for ZIP files and supported files starting with Comprehensive_Analysis
    candidates = []

    # Check for ZIP files
    for p in input_dir.iterdir():
        if p.is_file() and p.suffix.lower() == ".zip" and p.stem.startswith(COMPREHENSIVE_PREFIX):
            candidates.append(p)

    # Check for supported files (CSV, Excel)
    for p in input_dir.iterdir():
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS and p.stem.startswith(COMPREHENSIVE_PREFIX):
            candidates.append(p)

    if not candidates:
        return []

    # Oldest first: each processed export REPLACES the days it covers in the
    # history, so the newest export must be processed last to win (newest-first
    # let an older export overwrite a newer one's days).
    return sorted(candidates, key=lambda p: p.stat().st_mtime)


def write_manifest(output_root: Path, processed_files: List[dict]) -> None:
    manifest_path = output_root / "manifest.json"
    manifest = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "processed_files": processed_files,
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")


def run(input_dir: Path, output_root: Path, history_path: Path, logger: logging.Logger,
        recover_from: List[Path] = None) -> List[dict]:
    output_root.mkdir(parents=True, exist_ok=True)

    recovered = recover_stray_downloads([*(recover_from or []), input_dir], input_dir, logger)
    if recovered:
        logger.info(f"Recovered {len(recovered)} old/unprocessed export(s) into {input_dir}")

    files = discover_files(input_dir)
    processed_results: List[dict] = []

    if not files:
        logger.warning(f"No Comprehensive_Analysis files found in {input_dir}")
        return processed_results

    logger.info(f"Found {len(files)} file(s) to consider (including any just recovered)")

    ledger = load_processed_ledger(history_path)
    ledger_size_before = len(ledger)

    for file_path in files:
        logger.info(f"Considering: {file_path.name}")
        if file_path.suffix.lower() == ".zip":
            results = process_zip_file(file_path, output_root, history_path, logger, ledger)
            processed_results.extend(results)
        elif file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
            per_file_output_dir = output_root / "processed" / file_path.stem
            result = process_data_file(file_path, per_file_output_dir, history_path, logger, ledger)
            if result:
                processed_results.append(result)
        else:
            logger.info(f"Skipping unsupported file: {file_path}")

    if len(ledger) != ledger_size_before:
        save_processed_ledger(history_path, ledger)

    write_manifest(output_root, processed_results)
    return processed_results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Process downloaded ZIP/Excel/CSV files into analysis reports")
    parser.add_argument(
        "--source-dir",
        default=str(ANALYSIS_SOURCE_DIR),
        help=f"Folder to scan for new downloads (default: {ANALYSIS_SOURCE_DIR})",
    )
    parser.add_argument(
        "--output-dir",
        default=str(ANALYSIS_OUTPUT_DIR),
        help=f"Folder where processed files and reports will be written (default: {ANALYSIS_OUTPUT_DIR})",
    )
    parser.add_argument(
        "--history-file",
        default=str(ANALYSIS_HISTORY_FILE),
        help=f"One combined historical archive spreadsheet (default: {ANALYSIS_HISTORY_FILE})",
    )
    parser.add_argument(
        "--recover-from",
        nargs="*",
        default=[str(SMARTCARE_DOWNLOAD_DIR)],
        help="Additional folder(s) to sweep for old/unprocessed exports before analyzing "
             f"(default: SMARTCARE_DOWNLOAD_DIR = {SMARTCARE_DOWNLOAD_DIR})",
    )
    return parser.parse_args()


def ensure_directories():
    """Ensure all required directories exist."""
    ANALYSIS_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    ANALYSIS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ANALYSIS_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)


def main() -> int:
    args = parse_args()
    logger = setup_logger()

    # Ensure directories exist
    ensure_directories()

    source_dir = Path(args.source_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    history_file = Path(args.history_file).expanduser().resolve()
    recover_from = [Path(p).expanduser().resolve() for p in args.recover_from]

    logger.info(f"📂 Scanning source: {source_dir}")
    logger.info(f"📁 Writing reports to: {output_dir}")
    logger.info(f"📄 Using historical archive: {history_file}")
    logger.info(f"Also sweeping for old downloads in: {', '.join(str(p) for p in recover_from) or '(none)'}")

    run(source_dir, output_dir, history_file, logger, recover_from=recover_from)
    logger.info("✅ Processing complete")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
