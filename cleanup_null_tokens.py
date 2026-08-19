#!/usr/bin/env python3
"""
One-time cleanup: normalize Huawei's '/0' and 'NIL' null markers in the
existing output/csv/*.csv files to blank (NaN once re-read).

'/0' means a KPI ratio had zero attempts that hour (counter divide-by-zero),
NOT a 0% success rate - attempts were made and none failed vs. no attempts
happened at all are very different situations. 'NIL' means the counter
wasn't collected/supported for that cell. Both were stored as literal
strings, which forces the whole column to text dtype and corrupts every
downstream average/threshold check touching it.

backend/csv_loader.py's clean_null_tokens() now prevents this for new data
flowing through the pipeline (see read_csv_skip_metadata). This script
cleans what was already written to output/csv/ before that fix existed.
"""
import glob
import os

import pandas as pd

NULL_TOKENS = {'/0', 'NIL', 'NULL', 'N/A', '--'}
CSV_FOLDER = "output/csv"


def main():
    total_cells_cleaned = 0
    files_changed = 0
    for path in sorted(glob.glob(os.path.join(CSV_FOLDER, "*.csv"))):
        df = pd.read_csv(path, dtype=str, low_memory=False, keep_default_na=False)
        changed = False
        file_cleaned = 0
        for col in df.columns:
            mask = df[col].isin(NULL_TOKENS)
            n = int(mask.sum())
            if n:
                df.loc[mask, col] = ''
                changed = True
                file_cleaned += n
        if changed:
            df.to_csv(path, index=False)
            files_changed += 1
            total_cells_cleaned += file_cleaned
            print(f"Cleaned {os.path.basename(path)}: {file_cleaned} cells")

    print(f"\nDone. {files_changed} files changed, {total_cells_cleaned} cells normalized to blank/NaN.")


if __name__ == "__main__":
    main()
