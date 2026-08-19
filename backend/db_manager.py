# backend/db_manager.py
# !/usr/bin/env python3
"""
Libyana NPM - SQLite Database Manager
Database stored in FTP folder to keep project clean.
"""

import sqlite3
import pandas as pd
import os
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Database path in FTP folder
FTP_ROOT = r"C:\Users\user\Desktop\FTP files\Daily Report KPIs 2026 update"
DB_PATH = os.path.join(FTP_ROOT, "libyana_network.db")


class DatabaseManager:
    """
    SQLite database manager with automatic table creation,
    duplicate prevention, and data retrieval.
    Database is stored in the FTP folder.
    """

    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self._ensure_db_exists()
        self._create_all_tables()

    def _ensure_db_exists(self):
        """Ensure the database file exists."""
        # Create FTP folder if it doesn't exist
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)

        if not os.path.exists(self.db_path):
            logger.info(f"✅ Creating new database: {self.db_path}")
            open(self.db_path, 'w').close()

    def _get_connection(self):
        """Get a database connection."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _create_all_tables(self):
        """Create all tables if they don't exist."""
        conn = self._get_connection()
        cursor = conn.cursor()

        # Check if tables exist
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='site_summary'")
        if cursor.fetchone():
            logger.info("✅ Tables already exist")
            conn.close()
            return

        logger.info("📊 Creating database tables...")

        # ============================================================
        # CREATE ALL TABLES WITH UNIQUE CONSTRAINTS
        # ============================================================
        cursor.executescript("""
            -- 1. SITE SUMMARY - One row per day
            CREATE TABLE site_summary (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                total_physical_sites INTEGER DEFAULT 0,
                g2_physical_sites INTEGER DEFAULT 0,
                g2_gsm900_band INTEGER DEFAULT 0,
                g2_dcs1800_band INTEGER DEFAULT 0,
                g2_900_only INTEGER DEFAULT 0,
                g2_1800_only INTEGER DEFAULT 0,
                g3_physical_sites INTEGER DEFAULT 0,
                g3_u2100_band INTEGER DEFAULT 0,
                g3_u900_band INTEGER DEFAULT 0,
                g3_u2100_only INTEGER DEFAULT 0,
                g3_u900_only INTEGER DEFAULT 0,
                g4_physical_sites INTEGER DEFAULT 0,
                g4_l1800_f1_band INTEGER DEFAULT 0,
                g4_l1800_f2_band INTEGER DEFAULT 0,
                g4_l2100_band INTEGER DEFAULT 0,
                g4_l900_band INTEGER DEFAULT 0,
                g4_l700_band INTEGER DEFAULT 0,
                g4_l1800_only INTEGER DEFAULT 0,
                g4_l2100_only INTEGER DEFAULT 0,
                g4_l900_only INTEGER DEFAULT 0,
                g4_l700_only INTEGER DEFAULT 0,
                g2_only_sites INTEGER DEFAULT 0,
                g3_only_sites INTEGER DEFAULT 0,
                g4_only_sites INTEGER DEFAULT 0,
                g2_g3_sites INTEGER DEFAULT 0,
                g2_g4_sites INTEGER DEFAULT 0,
                g3_g4_sites INTEGER DEFAULT 0,
                g2_g3_g4_sites INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 2. SITE DETAIL - One row per site per day
            CREATE TABLE site_detail (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                site_name TEXT NOT NULL,
                g2_gsm900_band TEXT,
                g2_dcs1800_band TEXT,
                g3_u2100_band TEXT,
                g3_u900_band TEXT,
                g4_l1800_f1_band TEXT,
                g4_l1800_f2_band TEXT,
                g4_l2100_band TEXT,
                g4_l900_band TEXT,
                g4_l700_band TEXT,
                scenario TEXT,
                rat TEXT,
                sectors_number INTEGER,
                current_rat TEXT,
                day DATE NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(site_name, day)
            );

            -- 3. USER SUMMARY - One row per day
            CREATE TABLE user_summary (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                g2_cs_user INTEGER DEFAULT 0,
                g3_cs_user INTEGER DEFAULT 0,
                total_vlr_subscribers INTEGER DEFAULT 0,
                roaming_cs_almadar INTEGER DEFAULT 0,
                roaming_2g_ps_gb INTEGER DEFAULT 0,
                roaming_3g_ps_iu INTEGER DEFAULT 0,
                roaming_4g_ps_s1 INTEGER DEFAULT 0,
                g2_ps_user INTEGER DEFAULT 0,
                g3_ps_user INTEGER DEFAULT 0,
                g4_ps_user INTEGER DEFAULT 0,
                volte_user INTEGER DEFAULT 0,
                total_subscribers INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 4. TRAFFIC NETWORK TABLES - One row per day
            CREATE TABLE traffic_network_2g (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                g2_ps_traffic_gb REAL DEFAULT 0,
                g2_cs_traffic_erl REAL DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            CREATE TABLE traffic_network_3g (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                g3_ps_traffic_gb REAL DEFAULT 0,
                g3_cs_traffic_erl REAL DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            CREATE TABLE traffic_network_4g (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                g4_dl_traffic_gb REAL DEFAULT 0,
                g4_volte_traffic_erl REAL DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 5. NETWORK 2G - One row per day
            CREATE TABLE network_2g_nwbh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                whole_network TEXT,
                integrity TEXT,
                downlink_tbf_establishment_success_rate REAL,
                sdcch_congestion_rate REAL,
                tch_congestion_rate REAL,
                tch_drop_rate REAL,
                sdcch_drop_rate REAL,
                rr307_tch_availability REAL,
                call_setup_success_rate REAL,
                handover_success_rate REAL,
                ps_traffic_rlc_mb REAL,
                rca313_assignment_success_rate REAL,
                immediate_assignment_success_rate REAL,
                k3014_traffic_volume_on_tch_erl REAL,
                k3004_traffic_volume_on_sdcch_erl REAL,
                tl9114_average_throughput_downlink_gprs_kbps REAL,
                tl9333_average_throughput_downlink_egprs_kbps REAL,
                tl9232_average_throughput_uplink_egprs_kbps REAL,
                tl9014_average_throughput_uplink_gprs_kbps REAL,
                trx_availability_rate REAL,
                k3015_available_tchs REAL,
                k3016_configured_tchs REAL,
                interference_band_4_5 REAL,
                downlink_hqi_0_5 REAL,
                uplink_hqi_0_5 REAL,
                rf_resource_utilization_rate REAL,
                cm33_call_drops_traffic_channel REAL,
                k3013a_successful_tch_seizures REAL,
                ch323_successful_incoming_internal_handovers REAL,
                ch343_successful_incoming_external_handovers REAL,
                ch313_successful_outgoing_internal_handovers REAL,
                ch333_successful_outgoing_external_handovers REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            CREATE TABLE network_2g_nw_daily (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                whole_network TEXT,
                integrity TEXT,
                downlink_tbf_establishment_success_rate REAL,
                sdcch_congestion_rate REAL,
                tch_congestion_rate REAL,
                tch_drop_rate REAL,
                sdcch_drop_rate REAL,
                rr307_tch_availability REAL,
                call_setup_success_rate REAL,
                handover_success_rate REAL,
                ps_traffic_rlc_mb REAL,
                rca313_assignment_success_rate REAL,
                immediate_assignment_success_rate REAL,
                k3014_traffic_volume_on_tch_erl REAL,
                k3004_traffic_volume_on_sdcch_erl REAL,
                tl9114_average_throughput_downlink_gprs_kbps REAL,
                tl9333_average_throughput_downlink_egprs_kbps REAL,
                tl9232_average_throughput_uplink_egprs_kbps REAL,
                tl9014_average_throughput_uplink_gprs_kbps REAL,
                trx_availability_rate REAL,
                k3015_available_tchs REAL,
                k3016_configured_tchs REAL,
                interference_band_4_5 REAL,
                downlink_hqi_0_5 REAL,
                uplink_hqi_0_5 REAL,
                cm33_call_drops_traffic_channel REAL,
                k3013a_successful_tch_seizures REAL,
                ch323_successful_incoming_internal_handovers REAL,
                ch343_successful_incoming_external_handovers REAL,
                ch313_successful_outgoing_internal_handovers REAL,
                ch333_successful_outgoing_external_handovers REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 6. NETWORK 3G - One row per day
            CREATE TABLE network_3g_nwbh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                whole_network TEXT,
                integrity TEXT,
                ps_rab_setup_success_ratio REAL,
                ps_drop_ratio REAL,
                inter_rat_handover_success_rate_ps REAL,
                availability REAL,
                ps_traffic_ul_dl_gb REAL,
                hsdpa_throughput_per_user_kbps REAL,
                rrc_setup_success_ratio REAL,
                amr_rab_setup_success_rate REAL,
                call_setup_success_rate REAL,
                cs_call_drop_rate REAL,
                soft_handover_success_rate REAL,
                inter_rat_handover_success_rate_cs REAL,
                vs_hho_interfreqout_success_ratio REAL,
                intrafreq_hho_sr REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            CREATE TABLE network_3g_nw_daily (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                whole_network TEXT,
                integrity TEXT,
                rrc_setup_success_ratio REAL,
                amr_rab_setup_success_rate REAL,
                call_setup_success_rate REAL,
                cs_call_drop_rate REAL,
                soft_handover_success_rate REAL,
                inter_rat_handover_success_rate_cs REAL,
                availability REAL,
                vs_hho_interfreqout_success_ratio REAL,
                intrafreq_hho_sr REAL,
                hsdpa_throughput_per_user_kbps REAL,
                ps_rab_setup_success_ratio REAL,
                ps_drop_ratio REAL,
                inter_rat_handover_success_rate_ps REAL,
                ps_traffic_ul_dl_gb REAL,
                cs_traffic_erl REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 7. NETWORK 4G - One row per day
            CREATE TABLE network_4g_nwbh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                whole_network TEXT,
                integrity TEXT,
                rrc_setup_success_rate REAL,
                erab_setup_success_rate REAL,
                intra_ho_execute_success_rate REAL,
                service_drop_rate_all REAL,
                downlink_packet_loss REAL,
                uplink_packet_loss_rate REAL,
                user_downlink_average_throughput_mbps REAL,
                user_uplink_average_throughput_mbps REAL,
                dl_prb_utilizing_rate REAL,
                ul_prb_utilizing_rate REAL,
                csfb_preparation_success_rate REAL,
                l_ul_interference_avg_dbm REAL,
                radio_network_availability_rate REAL,
                l_cell_unavail_dur_sys REAL,
                l_cell_unavail_dur_manual REAL,
                l_traffic_user_avg REAL,
                l_traffic_user_max REAL,
                downlink_traffic_volume_gb REAL,
                uplink_traffic_volume_gb REAL,
                inter_freq_handover_success_rate REAL,
                s1_signal_connection_establishment_success_rate REAL,
                volte_traffic_volume_erl REAL,
                volte_setup_success_rate_zm REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            CREATE TABLE network_4g_nw_daily (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                whole_network TEXT,
                integrity TEXT,
                ul_traffic_volume_gb REAL,
                dl_traffic_volume_gb REAL,
                radio_network_availability_rate REAL,
                l_cell_unavail_dur_sys REAL,
                l_cell_unavail_dur_manual REAL,
                l_traffic_user_avg REAL,
                l_traffic_user_max REAL,
                dl_prb_utilizing_rate REAL,
                ul_prb_utilizing_rate REAL,
                volte_traffic_volume_erl REAL,
                volte_setup_success_rate_zm REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 8. CELL TABLES - One row per cell per day
            CREATE TABLE cell_2g_csbh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                gbsc TEXT,
                cell_ci INTEGER,
                cell_name TEXT NOT NULL,
                cell_index INTEGER,
                site_name TEXT,
                integrity TEXT,
                downlink_tbf_establishment_success_rate REAL,
                sdcch_congestion_rate REAL,
                tch_congestion_rate REAL,
                tch_drop_rate REAL,
                sdcch_drop_rate REAL,
                rr307_tch_availability REAL,
                call_setup_success_rate REAL,
                handover_success_rate REAL,
                ps_traffic_rlc_mb REAL,
                rca313_assignment_success_rate REAL,
                immediate_assignment_success_rate REAL,
                k3014_traffic_volume_on_tch_erl REAL,
                k3004_traffic_volume_on_sdcch_erl REAL,
                tl9114_average_throughput_downlink_gprs_kbps REAL,
                tl9333_average_throughput_downlink_egprs_kbps REAL,
                tl9232_average_throughput_uplink_egprs_kbps REAL,
                tl9014_average_throughput_uplink_gprs_kbps REAL,
                trx_availability_rate REAL,
                k3015_available_tchs REAL,
                k3016_configured_tchs REAL,
                interference_band_4_5 REAL,
                downlink_hqi_0_5 REAL,
                uplink_hqi_0_5 REAL,
                rf_resource_utilization_rate REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day, cell_name)
            );

            CREATE TABLE cell_3g_csbh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                rnc TEXT,
                cell_id INTEGER,
                cell_name TEXT NOT NULL,
                nodeb_name TEXT,
                integrity TEXT,
                rrc_setup_success_ratio REAL,
                amr_rab_setup_success_rate REAL,
                call_setup_success_rate REAL,
                cs_call_drop_rate REAL,
                soft_handover_success_rate REAL,
                inter_rat_handover_success_rate_cs REAL,
                availability REAL,
                vs_hho_interfreqout_success_ratio REAL,
                intrafreq_hho_sr REAL,
                hsdpa_throughput_per_user_kbps REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day, cell_name)
            );

            CREATE TABLE cell_4g_bh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                enodeb_name TEXT NOT NULL,
                cell_fdd_tdd_indication TEXT,
                cell_name TEXT,
                localcell_id INTEGER,
                enodeb_function_name TEXT,
                cell_active_state TEXT,
                subnet_name TEXT,
                integrity TEXT,
                rrc_setup_success_rate REAL,
                erab_setup_success_rate REAL,
                intra_ho_execute_success_rate REAL,
                service_drop_rate_all REAL,
                downlink_packet_loss REAL,
                uplink_packet_loss_rate REAL,
                user_downlink_average_throughput_mbps REAL,
                user_uplink_average_throughput_mbps REAL,
                dl_prb_utilizing_rate REAL,
                ul_prb_utilizing_rate REAL,
                csfb_preparation_success_rate REAL,
                l_ul_interference_avg_dbm REAL,
                radio_network_availability_rate REAL,
                l_cell_unavail_dur_sys REAL,
                l_cell_unavail_dur_manual REAL,
                l_traffic_user_avg REAL,
                l_traffic_user_max REAL,
                downlink_traffic_volume_gb REAL,
                uplink_traffic_volume_gb REAL,
                inter_freq_handover_success_rate REAL,
                s1_signal_connection_establishment_success_rate REAL,
                volte_traffic_volume_erl REAL,
                volte_setup_success_rate_zm REAL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day, enodeb_name)
            );

            -- 9. HEALTH HISTORY - One row per day
            CREATE TABLE health_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                overall_score REAL,
                lte_score REAL,
                umts_score REAL,
                gsm_score REAL,
                accessibility_score REAL,
                retainability_score REAL,
                mobility_score REAL,
                resource_utilization_score REAL,
                total_kpis INTEGER,
                passed_kpis INTEGER,
                failed_kpis INTEGER,
                alert_count INTEGER,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(day)
            );

            -- 10. ALERTS
            CREATE TABLE alerts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                day DATE NOT NULL,
                technology TEXT,
                kpi_name TEXT,
                value REAL,
                threshold REAL,
                severity TEXT,
                dimension TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            -- CREATE INDEXES
            CREATE INDEX idx_site_summary_day ON site_summary(day);
            CREATE INDEX idx_site_detail_day ON site_detail(day);
            CREATE INDEX idx_site_detail_site_name ON site_detail(site_name);
            CREATE INDEX idx_network_4g_nwbh_day ON network_4g_nwbh(day);
            CREATE INDEX idx_network_4g_nw_daily_day ON network_4g_nw_daily(day);
            CREATE INDEX idx_network_3g_nwbh_day ON network_3g_nwbh(day);
            CREATE INDEX idx_network_3g_nw_daily_day ON network_3g_nw_daily(day);
            CREATE INDEX idx_network_2g_nwbh_day ON network_2g_nwbh(day);
            CREATE INDEX idx_network_2g_nw_daily_day ON network_2g_nw_daily(day);
            CREATE INDEX idx_cell_2g_csbh_day ON cell_2g_csbh(day);
            CREATE INDEX idx_cell_2g_csbh_cell_name ON cell_2g_csbh(cell_name);
            CREATE INDEX idx_cell_3g_csbh_day ON cell_3g_csbh(day);
            CREATE INDEX idx_cell_3g_csbh_cell_name ON cell_3g_csbh(cell_name);
            CREATE INDEX idx_cell_4g_bh_day ON cell_4g_bh(day);
            CREATE INDEX idx_cell_4g_bh_enodeb_name ON cell_4g_bh(enodeb_name);
            CREATE INDEX idx_traffic_network_2g_day ON traffic_network_2g(day);
            CREATE INDEX idx_traffic_network_3g_day ON traffic_network_3g(day);
            CREATE INDEX idx_traffic_network_4g_day ON traffic_network_4g(day);
            CREATE INDEX idx_user_summary_day ON user_summary(day);
            CREATE INDEX idx_health_history_day ON health_history(day);
            CREATE INDEX idx_alerts_day ON alerts(day);
            CREATE INDEX idx_alerts_severity ON alerts(severity);
        """)

        conn.commit()
        conn.close()
        logger.info("✅ All database tables created successfully!")

    # ============================================================
    # GENERIC INSERT/UPDATE METHODS
    # ============================================================

    def insert_or_update(self, table_name: str, data: Dict, unique_columns: List[str]) -> Tuple[bool, str]:
        """
        Insert or update data with duplicate prevention.

        Returns:
            Tuple[bool, str]: (success, message)
            - (True, 'inserted') - New record inserted
            - (True, 'updated') - Existing record updated
            - (False, 'duplicate') - Duplicate detected
            - (False, 'error: ...') - Error occurred
        """
        if not data:
            return False, 'no_data'

        conn = self._get_connection()
        cursor = conn.cursor()

        try:
            # Build WHERE clause for unique check
            where_clause = " AND ".join([f"{col} = ?" for col in unique_columns])
            where_values = [data.get(col) for col in unique_columns]

            # Check if record exists
            cursor.execute(f"SELECT id FROM {table_name} WHERE {where_clause}", where_values)
            existing = cursor.fetchone()

            if existing:
                # Update existing record
                set_clause = ", ".join([f"{col} = ?" for col in data.keys()])
                values = list(data.values())
                cursor.execute(
                    f"UPDATE {table_name} SET {set_clause}, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                    values + [existing[0]]
                )
                conn.commit()
                logger.debug(f"🔄 Updated record in {table_name}: {data.get('day', 'unknown')}")
                return True, 'updated'
            else:
                # Insert new record
                columns = ", ".join(data.keys())
                placeholders = ", ".join(["?"] * len(data))
                cursor.execute(
                    f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})",
                    list(data.values())
                )
                conn.commit()
                logger.debug(f"✅ Inserted record in {table_name}: {data.get('day', 'unknown')}")
                return True, 'inserted'

        except sqlite3.IntegrityError as e:
            logger.warning(f"⚠️ Duplicate record in {table_name}: {e}")
            return False, f'duplicate: {e}'
        except Exception as e:
            logger.error(f"❌ Error in insert_or_update {table_name}: {e}")
            return False, f'error: {e}'
        finally:
            conn.close()

    def insert_or_update_dataframe(self, table_name: str, df: pd.DataFrame,
                                   unique_columns: List[str]) -> Dict:
        """
        Insert or update multiple records from a DataFrame.

        Returns:
            Dict: {'inserted': count, 'updated': count, 'duplicates': count, 'errors': count}
        """
        if df is None or df.empty:
            return {'inserted': 0, 'updated': 0, 'duplicates': 0, 'errors': 0}

        records = df.to_dict('records')

        inserted_count = 0
        updated_count = 0
        duplicate_count = 0
        error_count = 0

        for record in records:
            # Convert Date to day if needed
            if 'Date' in record and pd.notna(record['Date']):
                record['day'] = str(record['Date'])
            elif 'day' not in record and 'Date' in record:
                record['day'] = str(record['Date'])

            success, status = self.insert_or_update(table_name, record, unique_columns)

            if success:
                if status == 'inserted':
                    inserted_count += 1
                elif status == 'updated':
                    updated_count += 1
            else:
                if 'duplicate' in status:
                    duplicate_count += 1
                else:
                    error_count += 1

        return {
            'inserted': inserted_count,
            'updated': updated_count,
            'duplicates': duplicate_count,
            'errors': error_count
        }

    # ============================================================
    # RETRIEVAL METHODS
    # ============================================================

    def get_data(self, table_name: str, day: str = None,
                 limit: int = None, order_by: str = None) -> pd.DataFrame:
        """Get data from a table."""
        conn = self._get_connection()

        query = f"SELECT * FROM {table_name}"
        params = []

        if day:
            query += " WHERE day = ?"
            params.append(day)

        if order_by:
            query += f" ORDER BY {order_by}"

        if limit:
            query += " LIMIT ?"
            params.append(limit)

        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        return df

    def get_all_dates(self) -> List[str]:
        """Get all unique dates from site_summary."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT day FROM site_summary ORDER BY day DESC")
        dates = [row[0] for row in cursor.fetchall()]
        conn.close()
        return dates

    def get_latest_day(self, table_name: str) -> Optional[str]:
        """Get the latest day in a table."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(f"SELECT MAX(day) FROM {table_name}")
        result = cursor.fetchone()
        conn.close()
        return result[0] if result and result[0] else None

    def table_exists(self, table_name: str) -> bool:
        """Check if a table exists."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
        result = cursor.fetchone()
        conn.close()
        return result is not None

    def get_table_count(self, table_name: str) -> int:
        """Get row count for a table."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else 0

    def remove_duplicates(self, table_name: str, unique_columns: List[str]) -> int:
        """Remove duplicate records from a table."""
        conn = self._get_connection()
        cursor = conn.cursor()

        group_cols = ", ".join(unique_columns)

        delete_query = f"""
            DELETE FROM {table_name}
            WHERE id NOT IN (
                SELECT MIN(id)
                FROM {table_name}
                GROUP BY {group_cols}
            )
        """

        cursor.execute(delete_query)
        deleted_count = cursor.rowcount
        conn.commit()
        conn.close()

        if deleted_count > 0:
            logger.info(f"🗑️ Removed {deleted_count} duplicate rows from {table_name}")

        return deleted_count

    # ============================================================
    # SPECIFIC UPDATE METHODS
    # ============================================================

    def update_site_summary(self, row_dict: Dict) -> bool:
        """Update site summary for a day."""
        if not row_dict or 'day' not in row_dict:
            return False

        data = {
            'day': row_dict.get('day'),
            'total_physical_sites': row_dict.get('Total Physical Sites (2G+3G+4G)', 0),
            'g2_physical_sites': row_dict.get('2G physical sites', 0),
            'g2_gsm900_band': row_dict.get('2G GSM900 Band', 0),
            'g2_dcs1800_band': row_dict.get('2G DCS1800 Band', 0),
            'g2_900_only': row_dict.get('2G 900 only', 0),
            'g2_1800_only': row_dict.get('2G 1800 only', 0),
            'g3_physical_sites': row_dict.get('3G physical sites', 0),
            'g3_u2100_band': row_dict.get('3G U2100 Band', 0),
            'g3_u900_band': row_dict.get('3G U900 Band', 0),
            'g3_u2100_only': row_dict.get('3G U2100 only', 0),
            'g3_u900_only': row_dict.get('3G U900 only', 0),
            'g4_physical_sites': row_dict.get('4G physical sites', 0),
            'g4_l1800_f1_band': row_dict.get('4G L1800 F1 Band', 0),
            'g4_l1800_f2_band': row_dict.get('4G L1800 F2 Band', 0),
            'g4_l2100_band': row_dict.get('4G L2100 Band', 0),
            'g4_l900_band': row_dict.get('4G L900 Band', 0),
            'g4_l700_band': row_dict.get('4G L700 Band', 0),
            'g4_l1800_only': row_dict.get('4G L1800 only', 0),
            'g4_l2100_only': row_dict.get('4G L2100 only', 0),
            'g4_l900_only': row_dict.get('4G L900 only', 0),
            'g4_l700_only': row_dict.get('4G L700 only', 0),
            'g2_only_sites': row_dict.get('2G only sites', 0),
            'g3_only_sites': row_dict.get('3G only sites', 0),
            'g4_only_sites': row_dict.get('4G only sites', 0),
            'g2_g3_sites': row_dict.get('2G+3G sites', 0),
            'g2_g4_sites': row_dict.get('2G+4G sites', 0),
            'g3_g4_sites': row_dict.get('3G+4G sites', 0),
            'g2_g3_g4_sites': row_dict.get('2G+3G+4G sites', 0),
        }

        success, _ = self.insert_or_update('site_summary', data, ['day'])
        return success

    def update_user_summary(self, row_dict: Dict) -> bool:
        """Update user summary for a day."""
        if not row_dict or 'Date' not in row_dict:
            return False

        data = {
            'day': row_dict.get('Date'),
            'g2_cs_user': row_dict.get('2G CS user', 0),
            'g3_cs_user': row_dict.get('3G CS user', 0),
            'total_vlr_subscribers': row_dict.get('Total VLR Subscribers', 0),
            'roaming_cs_almadar': row_dict.get('Roaming CS (Almadar)', 0),
            'roaming_2g_ps_gb': row_dict.get('Roaming 2G PS (Gb)', 0),
            'roaming_3g_ps_iu': row_dict.get('Roaming 3G PS (Iu)', 0),
            'roaming_4g_ps_s1': row_dict.get('Roaming 4G PS (S1)', 0),
            'g2_ps_user': row_dict.get('2G PS user', 0),
            'g3_ps_user': row_dict.get('3G PS user', 0),
            'g4_ps_user': row_dict.get('4G PS user', 0),
            'volte_user': row_dict.get('VoLTE user', 0),
            'total_subscribers': row_dict.get('Total Subscribers', 0),
        }

        success, _ = self.insert_or_update('user_summary', data, ['day'])
        return success

    def update_traffic_network(self, tech: str, row_dict: Dict) -> bool:
        """Update traffic network data for a technology."""
        if not row_dict or 'Date' not in row_dict:
            return False

        table_map = {
            '2G': 'traffic_network_2g',
            '3G': 'traffic_network_3g',
            '4G': 'traffic_network_4g',
        }

        table = table_map.get(tech)
        if not table:
            return False

        data = {'day': row_dict.get('Date')}

        if tech == '2G':
            data['g2_ps_traffic_gb'] = row_dict.get('2G PS Traffic (GB)', 0)
            data['g2_cs_traffic_erl'] = row_dict.get('2G CS Traffic (Erl)', 0)
        elif tech == '3G':
            data['g3_ps_traffic_gb'] = row_dict.get('3G PS Traffic (GB)', 0)
            data['g3_cs_traffic_erl'] = row_dict.get('3G CS Traffic (Erl)', 0)
        elif tech == '4G':
            data['g4_dl_traffic_gb'] = row_dict.get('4G DL Traffic (GB)', 0)
            data['g4_volte_traffic_erl'] = row_dict.get('4G VoLTE Traffic (Erl)', 0)

        success, _ = self.insert_or_update(table, data, ['day'])
        return success

    def update_network_kpi(self, tech: str, kpi_type: str, row_dict: Dict) -> bool:
        """Update network KPI data for a technology."""
        if not row_dict or 'Date' not in row_dict:
            return False

        table_map = {
            ('2G', 'NWBH'): 'network_2g_nwbh',
            ('2G', 'NW_Daily'): 'network_2g_nw_daily',
            ('3G', 'NWBH'): 'network_3g_nwbh',
            ('3G', 'NW_Daily'): 'network_3g_nw_daily',
            ('4G', 'NWBH'): 'network_4g_nwbh',
            ('4G', 'NW_Daily'): 'network_4g_nw_daily',
        }

        table = table_map.get((tech, kpi_type))
        if not table:
            return False

        data = {'day': row_dict.get('Date')}

        # Add all columns from row_dict
        exclude_cols = ['Date', 'Whole Network', 'Integrity']
        for col, value in row_dict.items():
            if col not in exclude_cols and value is not None:
                db_col = col.lower().replace(' ', '_').replace('(', '').replace(')', '').replace('%', 'pct').replace(
                    '/', '_')
                db_col = ''.join(c for c in db_col if c.isalnum() or c == '_')
                data[db_col] = value

        if 'Whole Network' in row_dict:
            data['whole_network'] = row_dict.get('Whole Network')
        if 'Integrity' in row_dict:
            data['integrity'] = row_dict.get('Integrity')

        success, _ = self.insert_or_update(table, data, ['day'])
        return success

    def update_cell_kpi(self, sheet: str, row_dict: Dict) -> bool:
        """Update cell KPI data."""
        if not row_dict or 'Date' not in row_dict:
            return False

        table_map = {
            '2G_Cell_CSBH': 'cell_2g_csbh',
            '3G_Cell_CSBH': 'cell_3g_csbh',
            '4G_Cell_BH': 'cell_4g_bh',
        }

        table = table_map.get(sheet)
        if not table:
            return False

        data = {'day': row_dict.get('Date')}

        # Determine unique column
        if sheet == '2G_Cell_CSBH':
            unique_cols = ['day', 'cell_name']
            data['cell_name'] = row_dict.get('Cell Name', '')
        elif sheet == '3G_Cell_CSBH':
            unique_cols = ['day', 'cell_name']
            data['cell_name'] = row_dict.get('Cell Name', '')
        elif sheet == '4G_Cell_BH':
            unique_cols = ['day', 'enodeb_name']
            data['enodeb_name'] = row_dict.get('eNodeB Name', '')

        # Add all columns
        for col, value in row_dict.items():
            if col not in ['Date'] and value is not None:
                db_col = col.lower().replace(' ', '_').replace('(', '').replace(')', '').replace('%', 'pct').replace(
                    '/', '_')
                db_col = ''.join(c for c in db_col if c.isalnum() or c == '_')
                data[db_col] = value

        success, _ = self.insert_or_update(table, data, unique_cols)
        return success

    def get_table_info(self) -> Dict:
        """Get information about all tables."""
        conn = self._get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = cursor.fetchall()

        info = {}
        for table in tables:
            table_name = table[0]
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]
            cursor.execute(f"SELECT MAX(day) FROM {table_name}")
            latest = cursor.fetchone()[0]
            info[table_name] = {
                'count': count,
                'latest_day': latest if latest else None
            }

        conn.close()
        return info


# ============================================================
# SINGLETON INSTANCE
# ============================================================

_db_instance = None


def get_db() -> DatabaseManager:
    """Get the singleton database instance."""
    global _db_instance
    if _db_instance is None:
        _db_instance = DatabaseManager()
    return _db_instance