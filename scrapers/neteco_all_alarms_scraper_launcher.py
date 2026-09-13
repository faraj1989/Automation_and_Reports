"""Stable entry point for the NetEco all-alarms scraper - copied verbatim
(2026-09-10) from the sibling NOC Automation Suite (Automation-master),
then wired up to run in this project, same as scrapers/mae_scraper.py."""
from pathlib import Path
import runpy

runpy.run_path(str(Path(__file__).with_name("neteco_continuous all alrams.py")), run_name="__main__")
