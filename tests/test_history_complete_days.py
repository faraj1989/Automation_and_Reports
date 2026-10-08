"""
Tests for CSVHistoryManager._upsert_complete_days - the rule that KPI history
holds only complete business days, and a later export corrects a stored day.

Run:  .venv\\Scripts\\python.exe -m unittest discover -s tests -v
"""

import os
import sys
import shutil
import tempfile
import unittest
from datetime import date
from unittest import mock

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend import csv_history_manager as chm  # noqa: E402
from backend.history_integrity import check_history_integrity  # noqa: E402


def nwbh(rows):
    """rows: list of (date, volte_erl) -> a 4G-NWBH-shaped export frame."""
    return pd.DataFrame({
        'Date': [d for d, _ in rows],
        'Whole Network': 'Whole Network',
        'Integrity': '100%',
        'VoLTE Traffic Volume (Erl)': [v for _, v in rows],
    })


class CompleteDayHistoryTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.hm = chm.CSVHistoryManager(output_folder=self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_on(self, day, df, sheet='4G_NWBH', source='test-export'):
        """Ingest df as if the pipeline ran on `day`."""
        with mock.patch.object(chm, '_today', return_value=day):
            return self.hm._upsert_complete_days(sheet, df, ['Date', 'Whole Network'], source=source)

    def stored(self, sheet='4G_NWBH'):
        return pd.read_csv(os.path.join(self.tmp, 'csv', f'{sheet}.csv')).set_index('Date')

    def test_partial_today_row_is_not_stored(self):
        # 9/30 06:00 run: export holds complete 9/29 plus a partial 9/30 (00:00 busy hour)
        self.run_on(date(2026, 9, 30), nwbh([('2026-09-29', 600.0), ('2026-09-30 00:00', 380.0)]))
        s = self.stored()
        self.assertIn('2026-09-29', s.index)
        self.assertNotIn('2026-09-30', s.index)

    def test_later_export_corrects_stored_day(self):
        self.run_on(date(2026, 9, 29), nwbh([('2026-09-28', 385.4379)]))
        self.run_on(date(2026, 9, 30), nwbh([('2026-09-28', 604.4963), ('2026-09-29', 590.0)]))
        s = self.stored()
        self.assertAlmostEqual(s.at['2026-09-28', 'VoLTE Traffic Volume (Erl)'], 604.4963)
        self.assertEqual(len(s), 2)  # replaced, not duplicated

    def test_correction_is_audited(self):
        self.run_on(date(2026, 9, 29), nwbh([('2026-09-28', 385.4379)]), source='export-A')
        self.run_on(date(2026, 9, 30), nwbh([('2026-09-28', 604.4963)]), source='export-B')
        audit = pd.read_csv(self.hm.audit_log_path)
        self.assertEqual(len(audit), 1)
        row = audit.iloc[0]
        self.assertEqual(row['Column'], 'VoLTE Traffic Volume (Erl)')
        self.assertEqual(row['Key'], '2026-09-28 | Whole Network')
        self.assertAlmostEqual(float(row['Old Value']), 385.4379)
        self.assertAlmostEqual(float(row['New Value']), 604.4963)
        self.assertEqual(row['Source'], 'export-B')

    def test_identical_reexport_writes_no_audit(self):
        self.run_on(date(2026, 9, 29), nwbh([('2026-09-28', 604.4963)]))
        self.run_on(date(2026, 9, 30), nwbh([('2026-09-28', 604.4963)]))
        self.assertFalse(os.path.exists(self.hm.audit_log_path))

    def test_nil_stays_nan_and_is_not_a_change(self):
        # Huawei NIL = no valid samples -> NaN on both sides must not be logged as a revision
        self.run_on(date(2026, 9, 29), nwbh([('2026-09-28', float('nan'))]))
        self.run_on(date(2026, 9, 30), nwbh([('2026-09-28', float('nan'))]))
        self.assertTrue(pd.isna(self.stored().at['2026-09-28', 'VoLTE Traffic Volume (Erl)']))
        self.assertFalse(os.path.exists(self.hm.audit_log_path))

    def test_previously_stored_partial_row_is_purged(self):
        # History written by the old first-write-wins code still holds a partial 9/29
        os.makedirs(os.path.join(self.tmp, 'csv'), exist_ok=True)
        nwbh([('2026-09-28', 600.0), ('2026-09-29', 380.0)]).to_csv(
            os.path.join(self.tmp, 'csv', '4G_NWBH.csv'), index=False)
        self.run_on(date(2026, 9, 29), nwbh([('2026-09-28', 604.0), ('2026-09-29 00:00', 381.0)]))
        s = self.stored()
        self.assertEqual(list(s.index), ['2026-09-28'])

    def test_integrity_check_flags_future_and_duplicate_rows(self):
        csv_dir = os.path.join(self.tmp, 'csv')
        os.makedirs(csv_dir, exist_ok=True)
        nwbh([('2026-09-28', 1.0), ('2026-09-28', 2.0), ('2026-09-29', 3.0)]).to_csv(
            os.path.join(csv_dir, '4G_NWBH.csv'), index=False)
        res = check_history_integrity(csv_dir, today=date(2026, 9, 29)).set_index('Feed')
        self.assertEqual(res.at['4G_NWBH', 'Status'], 'FAIL')
        self.assertEqual(res.at['4G_NWBH', 'Future Rows'], 1)
        self.assertEqual(res.at['4G_NWBH', 'Duplicate Keys'], 1)

    def test_integrity_check_passes_clean_history(self):
        csv_dir = os.path.join(self.tmp, 'csv')
        os.makedirs(csv_dir, exist_ok=True)
        days = pd.date_range('2026-09-01', '2026-09-28').strftime('%Y-%m-%d')
        nwbh([(d, 1.0) for d in days]).to_csv(os.path.join(csv_dir, '4G_NWBH.csv'), index=False)
        res = check_history_integrity(csv_dir, today=date(2026, 9, 29)).set_index('Feed')
        self.assertEqual(res.at['4G_NWBH', 'Status'], 'PASS')


if __name__ == '__main__':
    unittest.main()
