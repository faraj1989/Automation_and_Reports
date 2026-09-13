"""SmartCare CEM analysis entry point.

Processes whatever SmartCare CEM exports are already on disk (in
SMARTCARE_OUTPUT_DIR, plus recovering any stray/unprocessed ones left in
SMARTCARE_DOWNLOAD_DIR) and folds them into the historical analysis archive.
Does not download anything itself - run the SmartCare CEM scraper
(scrapers/SmartCare CEM scraper  v11.py) separately first to fetch a new
export.
"""
import argparse
import importlib.util
import sys
from pathlib import Path

REPORTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(REPORTS_DIR.parent))

from project_config import load_env_file
from project_logging import setup_logger


def _load_module(filename: str, module_name: str):
    spec = importlib.util.spec_from_file_location(module_name, REPORTS_DIR / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_analysis(smartcare_output_dir=None, analysis_output_dir=None, history_file=None) -> dict:
    """Analyze existing SmartCare exports - plus recovering any stray/unprocessed
    ones from the download folder - into the historical archive. Returns the
    analysis results dict list and the history file path."""
    load_env_file()
    logger = setup_logger("smartcare-analysis")

    analysis = _load_module("download_analysis_pipeline.py", "download_analysis_pipeline")

    smartcare_output_dir = Path(smartcare_output_dir or analysis.SMARTCARE_OUTPUT_DIR)
    analysis_output_dir = Path(analysis_output_dir or analysis.ANALYSIS_OUTPUT_DIR)
    history_file = Path(history_file or analysis.ANALYSIS_HISTORY_FILE)
    download_dir = Path(analysis.SMARTCARE_DOWNLOAD_DIR)

    logger.info("=" * 60)
    logger.info(f"SmartCare output dir: {smartcare_output_dir}")
    logger.info(f"Analysis output dir: {analysis_output_dir}")
    logger.info(f"History file: {history_file}")
    logger.info("=" * 60)

    logger.info("Analyzing existing SmartCare exports and updating the historical archive...")
    results = analysis.run(
        smartcare_output_dir, analysis_output_dir, history_file, logger,
        recover_from=[download_dir],
    )

    logger.info(f"Done. {len(results)} export(s) newly folded into {history_file}")
    return {"processed": results, "history_file": str(history_file)}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Process existing SmartCare CEM exports - plus any older unprocessed ones "
                    "sitting in the download folder - into the historical analysis archive. "
                    "Does not download a new export."
    )
    parser.add_argument("--smartcare-output-dir", default=None)
    parser.add_argument("--analysis-output-dir", default=None)
    parser.add_argument("--history-file", default=None)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    run_analysis(
        smartcare_output_dir=args.smartcare_output_dir,
        analysis_output_dir=args.analysis_output_dir,
        history_file=args.history_file,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
