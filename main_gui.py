#!/usr/bin/env python3
"""
Libyana NPM - Main GUI
User interface for FTP downloader, Site Summary, Site Detail, Network KPIs, Cell KPIs, and Traffic KPIs.
"""

import os
import sys
import glob
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog
from datetime import datetime
import pandas as pd

# Replace the HistoryManager import with CSVHistoryManager
from backend.csv_history_manager import CSVHistoryManager

# Import backend modules
from backend import (
    ConfigManager,
    SFTPDownloader,
    process_site_day,
    process_all_days,
    get_latest_day_folder,
    SITE_SUMMARY_HEADER
)
from backend.site_detail_processor import SITE_DETAIL_HEADER

# Try to import tkcalendar
try:
    from tkcalendar import Calendar
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False


class LibyanaNPMApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Libyana NPM - Network Performance Monitor")
        self.root.geometry("1200x850")
        self.root.resizable(True, True)

        self._log_to_console("=" * 60)
        self._log_to_console("🚀 APPLICATION STARTING")
        self._log_to_console(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self._log_to_console("=" * 60)

        self.config = ConfigManager()
        self._log_to_console("✅ ConfigManager loaded")

        # Initialize CSV History Manager
        self.history_mgr = CSVHistoryManager()
        self._log_to_console("✅ CSVHistoryManager initialized")

        # GUI variables
        self.host_var = tk.StringVar(value=self.config.get('host', '10.171.68.77'))
        self.port_var = tk.StringVar(value=self.config.get('port', '22'))
        self.username_var = tk.StringVar(value=self.config.get('username', 'ftpuser'))
        self.password_var = tk.StringVar(value=self.config.get('password', ''))
        self.remote_path_var = tk.StringVar(value=self.config.get('remote_path', '/ftproot/Daily Report KPIs'))
        self.local_root_var = tk.StringVar(
            value=self.config.get('local_root', r'C:\Users\user\Desktop\FTP files\Daily Report KPIs 2026 update'))

        self._log_to_console(f"📁 Local Root: {self.local_root_var.get()}")

        # Summary variables
        self.day_display_map = {}
        self.selected_day = tk.StringVar()
        self.last_result = None
        self.last_day = None

        # Detail variables
        self.detail_df = None

        # Network KPIs variables
        self.network_results = None
        self.network_day = None

        # Cell KPIs variables
        self.cell_results = None
        self.cell_day = None

        # Traffic KPIs variables
        self.traffic_results = None
        self.traffic_day = None

        # Log queue
        self.log_queue = queue.Queue()

        self._build_widgets()
        self._poll_log_queue()

        self._log_to_console("🔄 Refreshing day list...")
        self._refresh_day_list()
        # User KPIs variables
        self.user_results = None
        self.user_summary_df = None
        self.user_day = None

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _log_to_console(self, message):
        """Print to console with timestamp for debugging."""
        timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
        print(f"[{timestamp}] {message}")

    def _build_widgets(self):
        self._log_to_console("🏗️ Building GUI widgets...")

        # Notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Tab 1: FTP Downloader
        self.ftp_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.ftp_tab, text="📥 FTP Downloader")
        self._build_ftp_tab()

        # Tab 2: Site Summary
        self.summary_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_tab, text="📊 Site Summary")
        self._build_summary_tab()

        # Tab 3: Network KPIs
        self.network_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.network_tab, text="📈 Network KPIs")
        self._build_network_kpi_tab()

        # Tab 4: Cell KPIs
        self.cell_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.cell_tab, text="📊 Cell KPIs")
        self._build_cell_kpi_tab()

        # Tab 5: Traffic KPIs
        self.traffic_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.traffic_tab, text="📊 Traffic KPIs")
        self._build_traffic_kpi_tab()

        # Tab 6: User KPIs (NEW)
        self.user_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.user_tab, text="👥 User KPIs")
        self._build_user_kpi_tab()

        # Tab 7: Site Detail
        self.detail_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.detail_tab, text="📋 Site Detail")
        self._build_detail_tab()

        # Tab 8: Log Output
        self.log_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.log_tab, text="📋 Log Output")
        self._build_log_tab()

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, padx=10, pady=(0, 5))

        self._log_to_console("✅ GUI widgets built")

    # ---------- FTP Tab ----------
    def _build_ftp_tab(self):
        main_frame = ttk.Frame(self.ftp_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        cfg_frame = ttk.LabelFrame(main_frame, text="FTP Configuration", padding="10")
        cfg_frame.pack(fill=tk.X, pady=5)

        ttk.Label(cfg_frame, text="Host:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(cfg_frame, textvariable=self.host_var, width=40).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(cfg_frame, text="Port:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(cfg_frame, textvariable=self.port_var, width=10).grid(row=1, column=1, sticky=tk.W, padx=5)

        ttk.Label(cfg_frame, text="Username:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(cfg_frame, textvariable=self.username_var, width=40).grid(row=2, column=1, sticky=tk.W, padx=5)

        ttk.Label(cfg_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Entry(cfg_frame, textvariable=self.password_var, width=40, show="*").grid(row=3, column=1, sticky=tk.W, padx=5)

        ttk.Label(cfg_frame, text="Remote Path:").grid(row=4, column=0, sticky=tk.W, pady=2)
        ttk.Entry(cfg_frame, textvariable=self.remote_path_var, width=40).grid(row=4, column=1, sticky=tk.W, padx=5)

        ttk.Label(cfg_frame, text="Local Root:").grid(row=5, column=0, sticky=tk.W, pady=2)
        ttk.Entry(cfg_frame, textvariable=self.local_root_var, width=40).grid(row=5, column=1, sticky=tk.W, padx=5)
        ttk.Button(cfg_frame, text="Browse...", command=self._browse_local_root).grid(row=5, column=2, padx=5)

        btn_frame = ttk.Frame(cfg_frame)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="💾 Save Config", command=self._save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔌 Test Connection", command=self._test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📥 Download All", command=self._start_download).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📅 Download Specific Date", command=self._prompt_date_download).pack(side=tk.LEFT, padx=5)

    # ---------- Summary Tab ----------
    def _build_summary_tab(self):
        main_frame = ttk.Frame(self.summary_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        day_frame = ttk.LabelFrame(main_frame, text="Select Day", padding="10")
        day_frame.pack(fill=tk.X, pady=5)

        ttk.Label(day_frame, text="Available Days:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.day_combo = ttk.Combobox(day_frame, textvariable=self.selected_day, state="readonly", width=30)
        self.day_combo.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.day_combo.bind('<<ComboboxSelected>>', self._on_day_selected)

        self.day_info_label = ttk.Label(day_frame, text="", foreground='gray')
        self.day_info_label.grid(row=0, column=2, sticky=tk.W, padx=10)

        ttk.Button(day_frame, text="🔄 Refresh", command=self._refresh_day_list).grid(row=0, column=3, padx=5)

        btn_frame = ttk.Frame(day_frame)
        btn_frame.grid(row=1, column=0, columnspan=4, pady=10)
        ttk.Button(btn_frame, text="▶ Run Summary", command=self._run_summary).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📅 Process All Days", command=self._run_all_days).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save to Excel", command=self._save_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📋 Copy Result", command=self._copy_result).pack(side=tk.LEFT, padx=5)
        # NEW: Combine All to Excel button
        ttk.Button(btn_frame, text="📊 Combine All to Excel", command=self._combine_all_to_excel).pack(side=tk.LEFT, padx=5)

        result_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.tree_frame = ttk.Frame(result_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.tree_frame, columns=('value',), height=15)
        self.tree.heading('#0', text='KPI')
        self.tree.heading('value', text='Value')
        self.tree.column('#0', width=400)
        self.tree.column('value', width=150)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

    # ---------- Network KPIs Tab ----------
    def _build_network_kpi_tab(self):
        """Build the Network KPIs tab."""
        main_frame = ttk.Frame(self.network_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X, pady=5)

        ttk.Label(info_frame, text="Processes 2G, 3G, 4G whole network KPIs (Busy Hour & Daily).").pack(anchor=tk.W)
        ttk.Label(info_frame, text="Files: 2G_NWBH, 2G_NW_Daily, 3G_NWBH, 3G_NW_Daily, 4G_NWBH, 4G_NW_Daily",
                  foreground='gray').pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="🔍 Process Network KPIs", command=self._run_network_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save to Excel", command=self._save_network_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📋 Copy to Clipboard", command=self._copy_network_kpis).pack(side=tk.LEFT, padx=5)

        result_frame = ttk.LabelFrame(main_frame, text="Network KPI Results", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.network_tree = ttk.Treeview(
            tree_container,
            columns=('Sheet', 'Rows', 'Columns', 'Date Range'),
            show='headings',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=10
        )
        self.network_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar.config(command=self.network_tree.yview)
        h_scrollbar.config(command=self.network_tree.xview)

        self.network_tree.heading('Sheet', text='Sheet Name')
        self.network_tree.heading('Rows', text='Rows')
        self.network_tree.heading('Columns', text='Columns')
        self.network_tree.heading('Date Range', text='Date Range')

        self.network_tree.column('Sheet', width=150)
        self.network_tree.column('Rows', width=80)
        self.network_tree.column('Columns', width=80)
        self.network_tree.column('Date Range', width=200)

        self.network_status = tk.StringVar(value="Ready - Click 'Process Network KPIs'")
        status_label = ttk.Label(main_frame, textvariable=self.network_status, foreground='gray')
        status_label.pack(anchor=tk.W, pady=5)

    # ---------- Cell KPIs Tab ----------
    def _build_cell_kpi_tab(self):
        """Build the Cell KPIs tab."""
        main_frame = ttk.Frame(self.cell_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X, pady=5)

        ttk.Label(info_frame, text="Processes 2G, 3G, 4G cell-level busy hour KPIs (CSBH/BH).").pack(anchor=tk.W)
        ttk.Label(info_frame, text="Files: 2G_Cell_CSBH, 3G_Cell_CSBH, 4G_Cell_BH", foreground='gray').pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="🔍 Process Cell KPIs", command=self._run_cell_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save to Excel", command=self._save_cell_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📋 Copy to Clipboard", command=self._copy_cell_kpis).pack(side=tk.LEFT, padx=5)

        result_frame = ttk.LabelFrame(main_frame, text="Cell KPI Results", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.cell_tree = ttk.Treeview(
            tree_container,
            columns=('Sheet', 'Rows', 'Columns', 'Date Range'),
            show='headings',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=10
        )
        self.cell_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar.config(command=self.cell_tree.yview)
        h_scrollbar.config(command=self.cell_tree.xview)

        self.cell_tree.heading('Sheet', text='Sheet Name')
        self.cell_tree.heading('Rows', text='Rows')
        self.cell_tree.heading('Columns', text='Columns')
        self.cell_tree.heading('Date Range', text='Date Range')

        self.cell_tree.column('Sheet', width=150)
        self.cell_tree.column('Rows', width=80)
        self.cell_tree.column('Columns', width=80)
        self.cell_tree.column('Date Range', width=200)

        self.cell_status = tk.StringVar(value="Ready - Click 'Process Cell KPIs'")
        status_label = ttk.Label(main_frame, textvariable=self.cell_status, foreground='gray')
        status_label.pack(anchor=tk.W, pady=5)

    # ---------- Traffic KPIs Tab ----------
    def _build_traffic_kpi_tab(self):
        """Build the Traffic KPIs tab."""
        main_frame = ttk.Frame(self.traffic_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X, pady=5)

        ttk.Label(info_frame, text="Processes 2G, 3G, 4G traffic KPIs (per site and whole network).").pack(anchor=tk.W)
        ttk.Label(info_frame, text="Files: Traffic_2G, Traffic_3G, Traffic_4G (per site & whole network)",
                  foreground='gray').pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="🔍 Process Traffic KPIs", command=self._run_traffic_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save to Excel", command=self._save_traffic_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📋 Copy to Clipboard", command=self._copy_traffic_kpis).pack(side=tk.LEFT, padx=5)

        result_frame = ttk.LabelFrame(main_frame, text="Traffic KPI Results", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.traffic_tree = ttk.Treeview(
            tree_container,
            columns=('Type', 'Sheet', 'Rows', 'Columns', 'Date Range'),
            show='headings',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=12
        )
        self.traffic_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar.config(command=self.traffic_tree.yview)
        h_scrollbar.config(command=self.traffic_tree.xview)

        self.traffic_tree.heading('Type', text='Type')
        self.traffic_tree.heading('Sheet', text='Sheet Name')
        self.traffic_tree.heading('Rows', text='Rows')
        self.traffic_tree.heading('Columns', text='Columns')
        self.traffic_tree.heading('Date Range', text='Date Range')

        self.traffic_tree.column('Type', width=80)
        self.traffic_tree.column('Sheet', width=150)
        self.traffic_tree.column('Rows', width=80)
        self.traffic_tree.column('Columns', width=80)
        self.traffic_tree.column('Date Range', width=200)

        self.traffic_status = tk.StringVar(value="Ready - Click 'Process Traffic KPIs'")
        status_label = ttk.Label(main_frame, textvariable=self.traffic_status, foreground='gray')
        status_label.pack(anchor=tk.W, pady=5)

    # ---------- User KPIs Tab ----------
    def _build_user_kpi_tab(self):
        """Build the User KPIs tab."""
        main_frame = ttk.Frame(self.user_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X, pady=5)

        ttk.Label(info_frame, text="Processes 5 user KPI files: CS Roaming, CS Subscribers, PS Roaming, PS Subscribers, VoLTE.").pack(anchor=tk.W)
        ttk.Label(info_frame, text="Consolidates all user data into a unified summary table.", foreground='gray').pack(anchor=tk.W)

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="🔍 Process User KPIs", command=self._run_user_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save to Excel", command=self._save_user_kpis).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📋 Copy to Clipboard", command=self._copy_user_kpis).pack(side=tk.LEFT, padx=5)

        # Results frame with Treeview
        result_frame = ttk.LabelFrame(main_frame, text="User Summary Results", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.user_tree = ttk.Treeview(
            tree_container,
            columns=('Sheet', 'Rows', 'Columns', 'Date Range'),
            show='headings',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=10
        )
        self.user_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar.config(command=self.user_tree.yview)
        h_scrollbar.config(command=self.user_tree.xview)

        self.user_tree.heading('Sheet', text='Sheet Name')
        self.user_tree.heading('Rows', text='Rows')
        self.user_tree.heading('Columns', text='Columns')
        self.user_tree.heading('Date Range', text='Date Range')

        self.user_tree.column('Sheet', width=180)
        self.user_tree.column('Rows', width=80)
        self.user_tree.column('Columns', width=80)
        self.user_tree.column('Date Range', width=200)

        self.user_status = tk.StringVar(value="Ready - Click 'Process User KPIs'")
        status_label = ttk.Label(main_frame, textvariable=self.user_status, foreground='gray')
        status_label.pack(anchor=tk.W, pady=5)

        # Store results
        self.user_results = None
        self.user_summary_df = None
        self.user_day = None

    # ---------- Log Tab ----------
    def _build_log_tab(self):
        main_frame = ttk.Frame(self.log_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        log_container = ttk.Frame(main_frame)
        log_container.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_container, height=20, width=80, state='normal')
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.log_text.tag_config('error', foreground='red')
        self.log_text.tag_config('info', foreground='black')
        self.log_text.tag_config('success', foreground='green')
        self.log_text.tag_config('warning', foreground='orange')
        self.log_text.tag_config('debug', foreground='blue')

        btn_frame = ttk.Frame(log_container)
        btn_frame.pack(side=tk.RIGHT, padx=5, pady=5, anchor=tk.N)

        ttk.Button(btn_frame, text="📋 Copy Log", command=self._copy_log).pack(pady=2)
        ttk.Button(btn_frame, text="🗑 Clear Log", command=self._clear_log).pack(pady=2)

    # ---------- User KPI Methods ----------
    def _run_user_kpis(self):
        """Process user KPIs from the selected day."""
        self._log_to_console("🔍 Process User KPIs clicked")
        self.user_status.set("Processing...")
        self._log("▶ Processing user KPIs...", 'info')
        self._run_thread(self._do_process_user_kpis)

    def _do_process_user_kpis(self):
        """Run the actual user KPI processing."""
        try:
            from backend.user_kpi_processor import process_user_kpis, aggregate_user_data
            from backend.site_detail_processor import get_latest_available_day
            from backend.csv_history_manager import CSVHistoryManager

            local_root = self.local_root_var.get().strip()
            latest_folder = get_latest_available_day(local_root, log_callback=self._log)

            if not latest_folder:
                self._log("❌ No data available", 'error')
                self.user_status.set("Error - No data available")
                return

            self.user_day = os.path.basename(os.path.dirname(latest_folder))
            self._log(f"📁 Using day: {self.user_day}")

            # Process raw user KPIs
            raw_results = process_user_kpis(latest_folder, log_callback=self._log)

            if not raw_results:
                self._log("❌ No user KPI data found", 'error')
                self.user_status.set("Error - No data")
                return

            self.user_results = raw_results
            self.root.after(0, self._update_user_tree, raw_results)

            # Save raw user KPIs to CSV
            hm = CSVHistoryManager()
            hm.update_user_kpis(raw_results)
            self._log("✅ Raw user KPIs saved to CSV", 'success')

            # Aggregate into summary table
            summary_df = aggregate_user_data(raw_results)
            if summary_df is not None and not summary_df.empty:
                self.user_summary_df = summary_df
                self.root.after(0, self._update_user_summary_tree, summary_df)

                # Save summary to CSV
                hm.update_user_summary(summary_df)
                self._log(f"✅ User summary saved to CSV ({len(summary_df)} rows)", 'success')
            else:
                self._log("⚠️ No user summary data generated", 'warning')

            total_sheets = sum(1 for df in raw_results.values() if df is not None and not df.empty)
            self.user_status.set(f"Ready - {total_sheets} sheets processed for {self.user_day}")

        except Exception as e:
            self._log_to_console(f"❌ Error processing user KPIs: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.user_status.set("Error")

    def _update_user_tree(self, results):
        """Update the treeview with user KPI results."""
        for item in self.user_tree.get_children():
            self.user_tree.delete(item)

        total_rows = 0
        total_sheets = 0

        for sheet_name, df in results.items():
            if df is not None and not df.empty:
                rows = len(df)
                cols = len(df.columns)
                total_rows += rows
                total_sheets += 1
                date_range = ""
                if 'Date' in df.columns:
                    try:
                        dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                        if len(dates) > 0:
                            date_range = f"{min(dates)} to {max(dates)}"
                    except Exception:
                        date_range = "N/A"
                self.user_tree.insert('', 'end', values=(sheet_name, rows, cols, date_range))
            else:
                self.user_tree.insert('', 'end', values=(sheet_name, 'No data', '-', '-'))

        self.user_tree.insert('', 'end', values=('-' * 20, '-' * 20, '-' * 20, '-' * 20))
        self.user_tree.insert('', 'end', values=('TOTAL', total_rows, total_sheets, ''))

    def _update_user_summary_tree(self, summary_df):
        """Update a separate treeview or show summary in the existing one."""
        self._log_to_console(f"📊 User Summary: {len(summary_df)} rows, {len(summary_df.columns)} columns")
        self._log_to_console(f"   Columns: {summary_df.columns.tolist()}")

    def _save_user_kpis(self):
        """Save user KPI results to CSV/Excel."""
        if self.user_results is None:
            messagebox.showwarning("No Data", "Please process user KPIs first.")
            return

        try:
            from backend.csv_history_manager import CSVHistoryManager
            hm = CSVHistoryManager()
            hm.update_user_kpis(self.user_results)
            if self.user_summary_df is not None:
                hm.update_user_summary(self.user_summary_df)
            excel_path = hm.export_to_excel()
            self._log("✅ User KPIs saved to CSV and Excel", 'success')
            messagebox.showinfo("Saved", f"User KPIs saved to:\n{excel_path if excel_path else 'CSV files'}")
        except Exception as e:
            self._log(f"❌ Error saving: {e}", 'error')
            messagebox.showerror("Error", f"Failed to save: {e}")

    def _copy_user_kpis(self):
        """Copy user KPI summary to clipboard."""
        if self.user_summary_df is None:
            messagebox.showwarning("No Data", "Please process user KPIs first.")
            return

        text = f"User KPIs Summary - {self.user_day}\n"
        text += "=" * 60 + "\n\n"

        # Show summary table
        text += self.user_summary_df.to_string(index=False)
        text += "\n\n" + "=" * 60

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.user_status.set("Copied to clipboard!")

    # ---------- Site Detail Tab ----------
    def _build_detail_tab(self):
        """Build the Site Detail tab."""
        main_frame = ttk.Frame(self.detail_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X, pady=5)

        ttk.Label(info_frame, text="Generates per-site detail view from the latest available day.").pack(anchor=tk.W)
        ttk.Label(info_frame, text="Shows all bands, scenario, RAT, sectors, and current RAT for each site.",
                  foreground='gray').pack(anchor=tk.W)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="🔍 Generate Site Detail", command=self._run_site_detail).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save to Excel", command=self._save_detail_to_excel).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📋 Copy to Clipboard", command=self._copy_detail).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📊 Export to CSV", command=self._export_detail_csv).pack(side=tk.LEFT, padx=5)

        result_frame = ttk.LabelFrame(main_frame, text="Site Detail Results", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_container = ttk.Frame(result_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.detail_tree = ttk.Treeview(
            tree_container,
            columns=SITE_DETAIL_HEADER,
            show='headings',
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=20
        )
        self.detail_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar.config(command=self.detail_tree.yview)
        h_scrollbar.config(command=self.detail_tree.xview)

        column_widths = {
            'Site Name': 120,
            '2G GSM900 Band': 100,
            '2G DCS1800 Band': 100,
            '3G U2100 Band': 90,
            '3G U900 Band': 90,
            '4G L1800 F1 Band': 130,
            '4G L1800 F2 Band': 130,
            '4G L2100 Band': 120,
            '4G L900 Band': 120,
            '4G L700 Band': 120,
            'Scenario': 100,
            'RAT': 60,
            'Sectors Number': 90,
            'Current RAT': 250
        }

        for col in SITE_DETAIL_HEADER:
            self.detail_tree.heading(col, text=col)
            self.detail_tree.column(col, width=column_widths.get(col, 100), minwidth=50)

        self.detail_status = tk.StringVar(value="Ready - Click 'Generate Site Detail'")
        status_label = ttk.Label(main_frame, textvariable=self.detail_status, foreground='gray')
        status_label.pack(anchor=tk.W, pady=5)

    # ---------- Combine All to Excel ----------
    def _combine_all_to_excel(self):
        """Combine all CSV files from output/csv to Historical_Network_Data.xlsx in one go."""
        self._log_to_console("=" * 60)
        self._log_to_console("📊 COMBINE ALL CSV TO EXCEL")
        self._log_to_console("=" * 60)

        self.status_var.set("Combining CSV to Excel...")
        self._log("▶ Combining all CSV files to Historical_Network_Data.xlsx...", 'info')
        self._run_thread(self._do_combine_all_to_excel)

    def _do_combine_all_to_excel(self):
        """Run the actual combine operation in a thread."""
        try:
            from backend.csv_history_manager import CSVHistoryManager

            # Get the CSV folder path
            csv_folder = os.path.join("output", "csv")

            if not os.path.exists(csv_folder):
                self._log(f"❌ CSV folder not found: {csv_folder}", 'error')
                self.status_var.set("Error - CSV folder not found")
                return

            # Count CSV files
            csv_files = [f for f in os.listdir(csv_folder) if f.endswith('.csv')]
            self._log(f"📁 Found {len(csv_files)} CSV files in {csv_folder}")

            if not csv_files:
                self._log("❌ No CSV files found to combine", 'error')
                self.status_var.set("Error - No CSV files")
                return

            # Create history manager and export
            self._log("📊 Exporting all CSVs to Excel...")
            history_mgr = CSVHistoryManager()
            excel_path = history_mgr.export_to_excel()

            if excel_path and os.path.exists(excel_path):
                self._log(f"✅ Successfully combined {len(csv_files)} CSV files to:", 'success')
                self._log(f"   📁 {excel_path}", 'success')
                self.status_var.set(f"Ready - Excel saved to {excel_path}")

                # Show file size
                file_size = os.path.getsize(excel_path) / (1024 * 1024)
                self._log(f"   📊 File size: {file_size:.2f} MB", 'success')

                # Ask user if they want to open the file
                self.root.after(0, lambda: self._ask_open_excel(excel_path))
            else:
                self._log("❌ Failed to combine CSV files", 'error')
                self.status_var.set("Error - Export failed")

        except Exception as e:
            self._log_to_console(f"❌ Error combining CSV: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.status_var.set("Error")

    def _ask_open_excel(self, excel_path):
        """Ask user if they want to open the Excel file."""
        response = messagebox.askyesno(
            "Excel Created",
            f"✅ Historical_Network_Data.xlsx created successfully!\n\n"
            f"File: {excel_path}\n"
            f"Size: {os.path.getsize(excel_path) / (1024 * 1024):.2f} MB\n\n"
            f"Do you want to open it now?"
        )
        if response:
            try:
                os.startfile(excel_path)
                self._log(f"📂 Opened: {excel_path}")
            except Exception as e:
                self._log(f"❌ Could not open file: {e}", 'error')

    # ---------- Methods ----------
    def _browse_local_root(self):
        self._log_to_console("📂 Browse Local Root clicked")
        path = filedialog.askdirectory(title="Select Local Root Directory")
        if path:
            self._log_to_console(f"📂 Selected path: {path}")
            self.local_root_var.set(path)
            self._refresh_day_list()

    def _save_config(self):
        self._log_to_console("💾 Saving configuration...")
        self.config.set('host', self.host_var.get().strip())
        self.config.set('port', self.port_var.get().strip())
        self.config.set('username', self.username_var.get().strip())
        self.config.set('password', self.password_var.get())
        self.config.set('remote_path', self.remote_path_var.get().strip())
        self.config.set('local_root', self.local_root_var.get().strip())
        self.config.save()
        self._log_to_console("✅ Configuration saved")
        self._log("✅ Configuration saved.", 'success')

    def _test_connection(self):
        self._log_to_console("🔌 Test Connection clicked")
        self.status_var.set("Testing connection...")
        self._run_thread(self._do_test_connection)

    def _do_test_connection(self):
        self._log_to_console("🔌 Running connection test...")
        host = self.host_var.get().strip()
        port = self.port_var.get().strip()
        username = self.username_var.get().strip()
        password = self.password_var.get()
        remote_path = self.remote_path_var.get().strip()

        self._log_to_console(f"   Host: {host}:{port}")
        self._log_to_console(f"   Username: {username}")
        self._log_to_console(f"   Remote Path: {remote_path}")

        downloader = SFTPDownloader(host, port, username, password, remote_path, "", log_callback=self._log)
        try:
            if downloader.connect():
                self._log_to_console("✅ Connection successful")
                self._log("✅ Connection successful.", 'success')
                files = downloader.list_remote_files()
                self._log_to_console(f"   Found {len(files)} files in remote directory")
                downloader.disconnect()
            else:
                self._log_to_console("❌ Connection failed")
                self._log("❌ Connection failed.", 'error')
        except Exception as e:
            self._log_to_console(f"❌ Test error: {e}")
            self._log(f"❌ Test error: {e}", 'error')
        finally:
            self.status_var.set("Ready")

    def _start_download(self):
        self._log_to_console("📥 Download All clicked")
        self.status_var.set("Downloading...")
        self._run_thread(self._do_download)

    def _do_download(self):
        self._log_to_console("📥 Starting download process...")
        host = self.host_var.get().strip()
        port = self.port_var.get().strip()
        username = self.username_var.get().strip()
        password = self.password_var.get()
        remote_path = self.remote_path_var.get().strip()
        local_root = self.local_root_var.get().strip()

        self._log_to_console(f"   Local Root: {local_root}")

        downloader = SFTPDownloader(host, port, username, password, remote_path, local_root, log_callback=self._log)
        try:
            if downloader.connect():
                self._log_to_console("✅ Connected, starting download...")
                downloader.download_and_organize()
                downloader.disconnect()
                self._log_to_console("✅ Download process completed")
                self._log("✅ Download process completed.", 'success')
                self._refresh_day_list()
            else:
                self._log_to_console("❌ Download aborted - connection failed")
                self._log("❌ Download aborted.", 'error')
        except Exception as e:
            self._log_to_console(f"❌ Download error: {e}")
            self._log(f"❌ Download error: {e}", 'error')
        finally:
            self.status_var.set("Ready")

    def _prompt_date_download(self):
        self._log_to_console("📅 Download Specific Date clicked")
        if HAS_TKCALENDAR:
            self._show_calendar_dialog()
        else:
            self._show_manual_date_dialog()

    def _show_calendar_dialog(self):
        self._log_to_console("📅 Showing calendar dialog")
        top = tk.Toplevel(self.root)
        top.title("Select Date")
        top.geometry("300x250")
        top.grab_set()

        cal = Calendar(top, selectmode='day', year=2026, month=8, day=1)
        cal.pack(pady=10)

        def confirm():
            date_obj = cal.selection_get()
            date_str = date_obj.strftime("%Y-%m-%d")
            self._log_to_console(f"📅 Selected date: {date_str}")
            top.destroy()
            self.status_var.set(f"Downloading for {date_str}...")
            self._run_thread(lambda: self._do_download_date(date_str))

        btn_frame = ttk.Frame(top)
        btn_frame.pack(pady=5)
        ttk.Button(btn_frame, text="Confirm", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=top.destroy).pack(side=tk.LEFT, padx=5)

    def _show_manual_date_dialog(self):
        self._log_to_console("📅 Showing manual date entry")
        date_str = simpledialog.askstring(
            "Date",
            "Enter date (YYYY-MM-DD):",
            parent=self.root
        )
        if date_str:
            self._log_to_console(f"📅 Entered date: {date_str}")
            try:
                datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                self._log_to_console("❌ Invalid date format")
                messagebox.showerror("Invalid Date", "Please use format YYYY-MM-DD")
                return
            self.status_var.set(f"Downloading for {date_str}...")
            self._run_thread(lambda: self._do_download_date(date_str))

    def _do_download_date(self, date_str):
        self._log_to_console(f"📥 Downloading for specific date: {date_str}")
        host = self.host_var.get().strip()
        port = self.port_var.get().strip()
        username = self.username_var.get().strip()
        password = self.password_var.get()
        remote_path = self.remote_path_var.get().strip()
        local_root = self.local_root_var.get().strip()

        downloader = SFTPDownloader(host, port, username, password, remote_path, local_root, log_callback=self._log)
        try:
            if downloader.connect():
                self._log_to_console(f"✅ Connected, downloading for {date_str}...")
                downloader.download_and_organize(target_date=date_str)
                downloader.disconnect()
                self._log_to_console(f"✅ Download for {date_str} completed")
                self._log(f"✅ Download for {date_str} completed.", 'success')
                self._refresh_day_list()
            else:
                self._log_to_console("❌ Download aborted - connection failed")
                self._log("❌ Download aborted.", 'error')
        except Exception as e:
            self._log_to_console(f"❌ Download error: {e}")
            self._log(f"❌ Download error: {e}", 'error')
        finally:
            self.status_var.set("Ready")

    def _refresh_day_list(self):
        """Refresh the list of available days with unzipped data."""
        self._log_to_console("=" * 50)
        self._log_to_console("🔄 REFRESHING DAY LIST")

        local_root = self.local_root_var.get().strip()
        self._log_to_console(f"📁 Local Root: {local_root}")

        if not os.path.exists(local_root):
            self._log_to_console(f"❌ Local root does not exist: {local_root}")
            self.day_display_map = {}
            self.day_combo['values'] = []
            self.day_info_label.config(text="Path not found")
            self.selected_day.set("")
            return

        self._log_to_console(f"✅ Local root exists: {local_root}")
        self._log_to_console(f"📂 Scanning: {local_root}")

        folders = []
        for item in os.listdir(local_root):
            item_path = os.path.join(local_root, item)
            if os.path.isdir(item_path):
                self._log_to_console(f"   Found folder: {item}")
                try:
                    datetime.strptime(item, '%Y-%m-%d')
                    self._log_to_console(f"   ✅ Valid date format: {item}")
                    unzipped_path = os.path.join(item_path, 'unzipped')
                    if os.path.exists(unzipped_path):
                        csv_count = len(glob.glob(os.path.join(unzipped_path, '*.csv')))
                        self._log_to_console(f"   📄 CSV files in unzipped: {csv_count}")
                        if csv_count > 0:
                            folders.append((item, csv_count))
                            self._log_to_console(f"   ✅ Added folder: {item} ({csv_count} CSVs)")
                        else:
                            self._log_to_console(f"   ⚠️ No CSV files in {item}/unzipped")
                    else:
                        self._log_to_console(f"   ⚠️ No unzipped folder in {item}")
                except ValueError:
                    self._log_to_console(f"   ⚠️ Skipping {item} (not a date format)")
            else:
                self._log_to_console(f"   Skipping file: {item}")

        if not folders:
            self._log_to_console("❌ No valid folders with CSV data found")
            self.day_display_map = {}
            self.day_combo['values'] = []
            self.day_info_label.config(text="No unzipped data found")
            self.selected_day.set("")
            return

        folders.sort(reverse=True)
        self._log_to_console(f"📋 Found {len(folders)} folders with CSV data")

        self.day_display_map = {}
        for date_str, csv_count in folders:
            display = f"{date_str} ({csv_count} CSVs)"
            self.day_display_map[display] = date_str
            self._log_to_console(f"   Map: '{display}' -> '{date_str}'")

        display_values = list(self.day_display_map.keys())
        self.day_combo['values'] = display_values
        self._log_to_console(f"📋 Dropdown values: {display_values}")

        if display_values:
            self.day_combo.set(display_values[0])
            self._log_to_console(f"📌 Set dropdown to: '{display_values[0]}'")
            self._on_day_selected()

        self._log_to_console("🔄 REFRESH COMPLETE")
        self._log_to_console("=" * 50)

    def _on_day_selected(self, event=None):
        """Handle day selection from dropdown."""
        selected_display = self.day_combo.get()
        self._log_to_console(f"📋 Day selected (from dropdown): '{selected_display}'")

        if not selected_display:
            self._log_to_console("⚠️ No selection made")
            self.day_info_label.config(text="No selection")
            self.selected_day.set("")
            return

        if "CSVs" in selected_display:
            if selected_display not in self.day_display_map:
                self._log_to_console(f"❌ '{selected_display}' not in mapping")
                self._log_to_console(f"   Available keys: {list(self.day_display_map.keys())}")
                self.day_info_label.config(text="Invalid selection")
                self.selected_day.set("")
                return
            day = self.day_display_map[selected_display]
            self._log_to_console(f"✅ Mapped display '{selected_display}' -> date '{day}'")
        else:
            day = selected_display
            self._log_to_console(f"📋 Using date directly: '{day}'")

        self.selected_day.set(day)

        local_root = self.local_root_var.get().strip()
        unzipped_path = os.path.join(local_root, day, 'unzipped')
        self._log_to_console(f"📁 Checking path: {unzipped_path}")

        if os.path.exists(unzipped_path):
            csv_count = len(glob.glob(os.path.join(unzipped_path, '*.csv')))
            self._log_to_console(f"✅ Found {csv_count} CSV files")
            self.day_info_label.config(text=f"{day}: {csv_count} CSV files")
        else:
            self._log_to_console(f"❌ Folder not found: {unzipped_path}")
            self.day_info_label.config(text=f"Folder not found: {day}")

    # ---------- Site Summary Methods ----------
    def _run_summary(self):
        """Run the site summary for the selected day."""
        self._log_to_console("=" * 50)
        self._log_to_console("▶ RUN SUMMARY CLICKED")

        day = self.selected_day.get()
        self._log_to_console(f"📋 Selected date: '{day}'")

        if not day:
            self._log_to_console("❌ No day selected")
            messagebox.showwarning("No Day Selected", "Please select a day first.")
            return

        local_root = self.local_root_var.get().strip()
        unzipped_path = os.path.join(local_root, day, 'unzipped')
        self._log_to_console(f"📁 Unzipped path: {unzipped_path}")

        if not os.path.exists(unzipped_path):
            self._log_to_console(f"❌ Folder not found: {unzipped_path}")
            messagebox.showerror("Folder Not Found", f"Unzipped folder not found:\n{unzipped_path}")
            return

        csv_files = glob.glob(os.path.join(unzipped_path, '*.csv'))
        self._log_to_console(f"📄 Found {len(csv_files)} CSV files in folder")
        for f in csv_files[:5]:
            self._log_to_console(f"   - {os.path.basename(f)}")
        if len(csv_files) > 5:
            self._log_to_console(f"   ... and {len(csv_files) - 5} more")

        self.status_var.set(f"Processing {day}...")
        self._log(f"▶ Starting site summary for: {day}", 'info')
        self._run_thread(self._do_summary, day, unzipped_path)

    def _do_summary(self, day, unzipped_path):
        """Run the actual summary processing."""
        self._log_to_console(f"📊 Processing summary for {day}...")
        try:
            result = process_site_day(unzipped_path, log_callback=self._log)

            if result is None:
                self._log_to_console("❌ process_site_day returned None")
                self._log("❌ Site summary failed - no result", 'error')
                self.status_var.set("Error - No result")
                return

            result['day'] = day
            self.last_result = result
            self.last_day = day

            self._log_to_console(f"✅ Summary complete for {day}")
            self._log_to_console(
                f"   Results: 2G={result.get('2G physical sites', 0)}, 3G={result.get('3G physical sites', 0)}, 4G={result.get('4G physical sites', 0)}")
            self._log_to_console(f"   Total Physical Sites: {result.get('Total Physical Sites (2G+3G+4G)', 0)}")

            # Update the treeview in the main thread
            self.root.after(0, self._update_results, result)

            # Save to CSV using CSVHistoryManager
            self.history_mgr.update_site_row(result)
            self._log("✅ Site summary saved to CSV", 'success')

            self.status_var.set(f"Ready - {day} processed")
        except Exception as e:
            self._log_to_console(f"❌ Error during processing: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error during processing: {e}", 'error')
            self.status_var.set("Error")

    def _run_all_days(self):
        """Process all days from all folders."""
        self._log_to_console("=" * 60)
        self._log_to_console("📅 PROCESS ALL DAYS CLICKED")

        local_root = self.local_root_var.get().strip()
        if not os.path.exists(local_root):
            self._log_to_console(f"❌ Local root does not exist: {local_root}")
            messagebox.showerror("Folder Not Found", f"Local root not found:\n{local_root}")
            return

        self.status_var.set("Processing all days...")
        self._log(f"▶ Starting processing for all days", 'info')
        self._run_thread(self._do_process_all_days, local_root)

    def _do_process_all_days(self, local_root):
        """Run the actual processing for all days."""
        self._log_to_console(f"📊 Processing all days from: {local_root}")

        try:
            results = process_all_days(local_root, log_callback=self._log)

            if not results:
                self._log_to_console("❌ No results from processing all days")
                self._log("❌ No results found.", 'error')
                self.status_var.set("Ready - No results")
                return

            self._log_to_console(f"✅ Processed {len(results)} days")
            self._log(f"✅ Processed {len(results)} days", 'success')

            days = sorted(results.keys(), reverse=True)
            if days:
                latest_day = days[0]
                self.last_result = results[latest_day]
                self.last_day = latest_day
                self.root.after(0, self._update_results, self.last_result)
                self._log(f"📊 Showing latest: {latest_day}", 'info')

            # Save each day to CSV using CSVHistoryManager
            self._log_to_console("💾 Saving all days to CSV...")
            for day, result in results.items():
                result['day'] = day
                self.history_mgr.update_site_row(result)
            self._log(f"✅ Saved {len(results)} days to CSV", 'success')

            self.root.after(0, lambda: self._ask_save_all_days(results))

            self.status_var.set(f"Ready - Processed {len(results)} days")

        except Exception as e:
            self._log_to_console(f"❌ Error processing all days: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.status_var.set("Error")

    def _ask_save_all_days(self, results):
        """Ask user if they want to export all results to Excel."""
        if not results:
            return

        response = messagebox.askyesno(
            "Export to Excel",
            f"Processed {len(results)} days. Do you want to export all data to Excel?"
        )

        if response:
            self._log_to_console("📁 Exporting all data to Excel...")
            try:
                excel_path = self.history_mgr.export_to_excel()
                if excel_path:
                    self._log(f"✅ Exported to {excel_path}", 'success')
                    messagebox.showinfo("Export Complete", f"Data exported to:\n{excel_path}")
                else:
                    self._log("❌ No data to export", 'error')
            except Exception as e:
                self._log(f"❌ Export failed: {e}", 'error')
                messagebox.showerror("Export Failed", str(e))

    def _update_results(self, result):
        """Update the treeview with results."""
        self._log_to_console("📊 Updating results treeview...")
        for item in self.tree.get_children():
            self.tree.delete(item)

        groups = {
            '2G': ['2G physical sites', '2G GSM900 Band', '2G DCS1800 Band', '2G 900 only', '2G 1800 only'],
            '3G': ['3G physical sites', '3G U2100 Band', '3G U900 Band', '3G U2100 only', '3G U900 only'],
            '4G': ['4G physical sites', '4G L1800 F1 Band', '4G L1800 F2 Band', '4G L2100 Band',
                   '4G L900 Band', '4G L700 Band', '4G L1800 only', '4G L2100 only', '4G L900 only', '4G L700 only'],
            'Overlaps': ['Total Physical Sites (2G+3G+4G)', '2G only sites', '3G only sites', '4G only sites',
                         '2G+3G sites', '2G+4G sites', '3G+4G sites', '2G+3G+4G sites']
        }

        for group_name, keys in groups.items():
            group_id = self.tree.insert('', 'end', text=f'── {group_name} ──', values=('',))
            for key in keys:
                value = result.get(key, 0)
                self.tree.insert(group_id, 'end', text=f'  {key}', values=(f'{value:,}',))

        self._log_to_console("✅ Results treeview updated")

    def _save_to_excel(self):
        """Save the current result to Excel."""
        if self.last_result is None:
            messagebox.showwarning("No Result", "Please run the site summary first.")
            return

        try:
            # Update CSV first
            self.history_mgr.update_site_row(self.last_result)
            # Then export all to Excel (this will create Historical_Network_Data.xlsx in output/)
            excel_path = self.history_mgr.export_to_excel()
            if excel_path:
                self._log(f"✅ Data saved to {excel_path}", 'success')
                messagebox.showinfo("Saved", f"Data saved to {excel_path}")
            else:
                self._log("❌ No data to export", 'error')
        except Exception as e:
            self._log(f"❌ Error saving to Excel: {e}", 'error')
            messagebox.showerror("Save Error", f"Failed to save: {e}")

    def _copy_result(self):
        """Copy the result to clipboard."""
        self._log_to_console("📋 Copy Result clicked")
        if self.last_result is None:
            self._log_to_console("❌ No result to copy")
            messagebox.showwarning("No Result", "Please run the site summary first.")
            return

        text = f"Site Summary - {self.last_day}\n"
        text += "=" * 50 + "\n"
        for key, value in self.last_result.items():
            text += f"{key}: {value:,}\n"

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._log_to_console("✅ Results copied to clipboard")
        self.status_var.set("Results copied to clipboard!")

    # ---------- Network KPI Methods ----------
    def _run_network_kpis(self):
        """Process network KPIs from the selected day."""
        self._log_to_console("🔍 Process Network KPIs clicked")
        self.network_status.set("Processing...")
        self._log("▶ Processing network KPIs...", 'info')
        self._run_thread(self._do_process_network_kpis)

    def _do_process_network_kpis(self):
        """Run the actual network KPI processing."""
        try:
            from backend.network_kpi_processor import process_network_kpis
            from backend.site_detail_processor import get_latest_available_day

            local_root = self.local_root_var.get().strip()
            latest_folder = get_latest_available_day(local_root, log_callback=self._log)

            if not latest_folder:
                self._log("❌ No data available", 'error')
                self.network_status.set("Error - No data available")
                return

            self.network_day = os.path.basename(os.path.dirname(latest_folder))
            self._log(f"📁 Using day: {self.network_day}")

            results = process_network_kpis(latest_folder, log_callback=self._log)

            if not results:
                self._log("❌ No network KPI data found", 'error')
                self.network_status.set("Error - No data")
                return

            self.network_results = results
            self.root.after(0, self._update_network_tree, results)

            # Save to CSV using CSVHistoryManager
            self.history_mgr.update_network_kpis(results)
            self._log("✅ Network KPIs saved to CSV", 'success')

            total_rows = sum(len(df) for df in results.values() if df is not None and not df.empty)
            total_sheets = sum(1 for df in results.values() if df is not None and not df.empty)

            self.network_status.set(
                f"Ready - {total_sheets} sheets, {total_rows} rows processed for {self.network_day}")

        except Exception as e:
            self._log_to_console(f"❌ Error processing network KPIs: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.network_status.set("Error")

    def _update_network_tree(self, results):
        """Update the treeview with network KPI results."""
        for item in self.network_tree.get_children():
            self.network_tree.delete(item)

        total_rows = 0
        total_sheets = 0

        for sheet_name, df in results.items():
            if df is not None and not df.empty:
                rows = len(df)
                cols = len(df.columns)
                total_rows += rows
                total_sheets += 1
                date_range = ""
                if 'Date' in df.columns:
                    try:
                        dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                        if len(dates) > 0:
                            date_range = f"{min(dates)} to {max(dates)}"
                    except Exception:
                        date_range = "N/A"
                self.network_tree.insert('', 'end', values=(sheet_name, rows, cols, date_range))
            else:
                self.network_tree.insert('', 'end', values=(sheet_name, 'No data', '-', '-'))

        self.network_tree.insert('', 'end', values=('-' * 20, '-' * 20, '-' * 20, '-' * 20))
        self.network_tree.insert('', 'end', values=('TOTAL', total_rows, total_sheets, ''))

    def _save_network_kpis(self):
        """Save network KPI results to CSV/Excel."""
        if self.network_results is None:
            messagebox.showwarning("No Data", "Please process network KPIs first.")
            return

        try:
            self.history_mgr.update_network_kpis(self.network_results)
            excel_path = self.history_mgr.export_to_excel()
            self._log("✅ Network KPIs saved to CSV and Excel", 'success')
            messagebox.showinfo("Saved", f"Network KPIs saved to:\n{excel_path if excel_path else 'CSV files'}")
        except Exception as e:
            self._log(f"❌ Error saving: {e}", 'error')
            messagebox.showerror("Error", f"Failed to save: {e}")

    def _copy_network_kpis(self):
        """Copy network KPI summary to clipboard."""
        if self.network_results is None:
            messagebox.showwarning("No Data", "Please process network KPIs first.")
            return

        text = f"Network KPIs - {self.network_day}\n"
        text += "=" * 60 + "\n"

        for sheet_name, df in self.network_results.items():
            if df is not None and not df.empty:
                text += f"\n{sheet_name}:\n"
                text += f"  Rows: {len(df)}, Columns: {len(df.columns)}\n"
                if 'Date' in df.columns:
                    try:
                        dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                        text += f"  Date Range: {min(dates)} to {max(dates)}\n"
                    except Exception:
                        pass
                if 'Date' in df.columns and 'Whole Network' in df.columns:
                    text += f"  Sample dates: {df['Date'].head(3).tolist()}\n"
            else:
                text += f"\n{sheet_name}: No data\n"

        text += "\n" + "=" * 60
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.network_status.set("Copied to clipboard!")

    # ---------- Cell KPI Methods ----------
    def _run_cell_kpis(self):
        """Process cell KPIs from the selected day."""
        self._log_to_console("🔍 Process Cell KPIs clicked")
        self.cell_status.set("Processing...")
        self._log("▶ Processing cell KPIs...", 'info')
        self._run_thread(self._do_process_cell_kpis)

    def _do_process_cell_kpis(self):
        """Run the actual cell KPI processing."""
        try:
            from backend.cell_kpi_processor import process_cell_kpis
            from backend.site_detail_processor import get_latest_available_day

            local_root = self.local_root_var.get().strip()
            latest_folder = get_latest_available_day(local_root, log_callback=self._log)

            if not latest_folder:
                self._log("❌ No data available", 'error')
                self.cell_status.set("Error - No data available")
                return

            self.cell_day = os.path.basename(os.path.dirname(latest_folder))
            self._log(f"📁 Using day: {self.cell_day}")

            results = process_cell_kpis(latest_folder, log_callback=self._log)

            if not results:
                self._log("❌ No cell KPI data found", 'error')
                self.cell_status.set("Error - No data")
                return

            self.cell_results = results
            self.root.after(0, self._update_cell_tree, results)

            # Save to CSV using CSVHistoryManager
            self.history_mgr.update_cell_kpis(results)
            self._log("✅ Cell KPIs saved to CSV", 'success')

            total_rows = sum(len(df) for df in results.values() if df is not None and not df.empty)
            total_sheets = sum(1 for df in results.values() if df is not None and not df.empty)

            self.cell_status.set(f"Ready - {total_sheets} sheets, {total_rows} rows processed for {self.cell_day}")

        except Exception as e:
            self._log_to_console(f"❌ Error processing cell KPIs: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.cell_status.set("Error")

    def _update_cell_tree(self, results):
        """Update the treeview with cell KPI results."""
        for item in self.cell_tree.get_children():
            self.cell_tree.delete(item)

        total_rows = 0
        total_sheets = 0

        for sheet_name, df in results.items():
            if df is not None and not df.empty:
                rows = len(df)
                cols = len(df.columns)
                total_rows += rows
                total_sheets += 1
                date_range = ""
                if 'Date' in df.columns:
                    try:
                        dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                        if len(dates) > 0:
                            date_range = f"{min(dates)} to {max(dates)}"
                    except Exception:
                        date_range = "N/A"
                self.cell_tree.insert('', 'end', values=(sheet_name, rows, cols, date_range))
            else:
                self.cell_tree.insert('', 'end', values=(sheet_name, 'No data', '-', '-'))

        self.cell_tree.insert('', 'end', values=('-' * 20, '-' * 20, '-' * 20, '-' * 20))
        self.cell_tree.insert('', 'end', values=('TOTAL', total_rows, total_sheets, ''))

    def _save_cell_kpis(self):
        """Save cell KPI results to CSV/Excel."""
        if self.cell_results is None:
            messagebox.showwarning("No Data", "Please process cell KPIs first.")
            return

        try:
            self.history_mgr.update_cell_kpis(self.cell_results)
            excel_path = self.history_mgr.export_to_excel()
            self._log("✅ Cell KPIs saved to CSV and Excel", 'success')
            messagebox.showinfo("Saved", f"Cell KPIs saved to:\n{excel_path if excel_path else 'CSV files'}")
        except Exception as e:
            self._log(f"❌ Error saving: {e}", 'error')
            messagebox.showerror("Error", f"Failed to save: {e}")

    def _copy_cell_kpis(self):
        """Copy cell KPI summary to clipboard."""
        if self.cell_results is None:
            messagebox.showwarning("No Data", "Please process cell KPIs first.")
            return

        text = f"Cell KPIs - {self.cell_day}\n"
        text += "=" * 60 + "\n"

        for sheet_name, df in self.cell_results.items():
            if df is not None and not df.empty:
                text += f"\n{sheet_name}:\n"
                text += f"  Rows: {len(df)}, Columns: {len(df.columns)}\n"
                if 'Date' in df.columns:
                    try:
                        dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                        text += f"  Date Range: {min(dates)} to {max(dates)}\n"
                    except Exception:
                        pass
            else:
                text += f"\n{sheet_name}: No data\n"

        text += "\n" + "=" * 60
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.cell_status.set("Copied to clipboard!")

    # ---------- Traffic KPI Methods ----------
    def _run_traffic_kpis(self):
        """Process traffic KPIs from the selected day."""
        self._log_to_console("🔍 Process Traffic KPIs clicked")
        self.traffic_status.set("Processing...")
        self._log("▶ Processing traffic KPIs...", 'info')
        self._run_thread(self._do_process_traffic_kpis)

    def _do_process_traffic_kpis(self):
        """Run the actual traffic KPI processing."""
        try:
            from backend.traffic_kpi_processor import process_traffic_with_aggregation
            from backend.site_detail_processor import get_latest_available_day

            local_root = self.local_root_var.get().strip()
            latest_folder = get_latest_available_day(local_root, log_callback=self._log)

            if not latest_folder:
                self._log("❌ No data available", 'error')
                self.traffic_status.set("Error - No data available")
                return

            self.traffic_day = os.path.basename(os.path.dirname(latest_folder))
            self._log(f"📁 Using day: {self.traffic_day}")

            results = process_traffic_with_aggregation(latest_folder, log_callback=self._log)

            if not results:
                self._log("❌ No traffic KPI data found", 'error')
                self.traffic_status.set("Error - No data")
                return

            self.traffic_results = results
            self.root.after(0, self._update_traffic_tree, results)

            # Save to CSV using CSVHistoryManager
            self.history_mgr.update_traffic_kpis(results)
            self._log("✅ Traffic KPIs saved to CSV", 'success')

            total_rows = 0
            total_sheets = 0
            for sheet_type in ['per_site', 'whole_network']:
                if sheet_type in results:
                    for sheet_name, df in results[sheet_type].items():
                        if df is not None and not df.empty:
                            total_rows += len(df)
                            total_sheets += 1

            self.traffic_status.set(
                f"Ready - {total_sheets} sheets, {total_rows} rows processed for {self.traffic_day}")

        except Exception as e:
            self._log_to_console(f"❌ Error processing traffic KPIs: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.traffic_status.set("Error")

    def _update_traffic_tree(self, results):
        """Update the treeview with traffic KPI results."""
        for item in self.traffic_tree.get_children():
            self.traffic_tree.delete(item)

        total_rows = 0
        total_sheets = 0

        for sheet_type, sheets in results.items():
            type_label = "Per Site" if sheet_type == 'per_site' else "Whole Network"
            for sheet_name, df in sheets.items():
                if df is not None and not df.empty:
                    rows = len(df)
                    cols = len(df.columns)
                    total_rows += rows
                    total_sheets += 1
                    date_range = ""
                    if 'Date' in df.columns:
                        try:
                            dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                            if len(dates) > 0:
                                date_range = f"{min(dates)} to {max(dates)}"
                        except Exception:
                            date_range = "N/A"
                    self.traffic_tree.insert('', 'end', values=(type_label, sheet_name, rows, cols, date_range))
                else:
                    self.traffic_tree.insert('', 'end', values=(type_label, sheet_name, 'No data', '-', '-'))

        self.traffic_tree.insert('', 'end', values=('-' * 10, '-' * 10, '-' * 10, '-' * 10, '-' * 10))
        self.traffic_tree.insert('', 'end', values=('TOTAL', '', total_rows, total_sheets, ''))

    def _save_traffic_kpis(self):
        """Save traffic KPI results to CSV/Excel."""
        if self.traffic_results is None:
            messagebox.showwarning("No Data", "Please process traffic KPIs first.")
            return

        try:
            self.history_mgr.update_traffic_kpis(self.traffic_results)
            excel_path = self.history_mgr.export_to_excel()
            self._log("✅ Traffic KPIs saved to CSV and Excel", 'success')
            messagebox.showinfo("Saved", f"Traffic KPIs saved to:\n{excel_path if excel_path else 'CSV files'}")
        except Exception as e:
            self._log(f"❌ Error saving: {e}", 'error')
            messagebox.showerror("Error", f"Failed to save: {e}")

    def _copy_traffic_kpis(self):
        """Copy traffic KPI summary to clipboard."""
        if self.traffic_results is None:
            messagebox.showwarning("No Data", "Please process traffic KPIs first.")
            return

        text = f"Traffic KPIs - {self.traffic_day}\n"
        text += "=" * 60 + "\n"

        for sheet_type, sheets in self.traffic_results.items():
            type_label = "Per Site" if sheet_type == 'per_site' else "Whole Network"
            text += f"\n{type_label}:\n"
            for sheet_name, df in sheets.items():
                if df is not None and not df.empty:
                    text += f"  {sheet_name}: {len(df)} rows, {len(df.columns)} columns\n"
                    if 'Date' in df.columns:
                        try:
                            dates = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d').unique()
                            text += f"    Date Range: {min(dates)} to {max(dates)}\n"
                        except Exception:
                            pass
                else:
                    text += f"  {sheet_name}: No data\n"

        text += "\n" + "=" * 60
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.traffic_status.set("Copied to clipboard!")

    # ---------- Site Detail Methods ----------
    def _run_site_detail(self):
        """Generate site detail from the latest available day."""
        self._log_to_console("🔍 Generate Site Detail clicked")
        self.detail_status.set("Generating...")
        self._log("▶ Generating site detail...", 'info')
        self._run_thread(self._do_generate_detail)

    def _do_generate_detail(self):
        """Run the actual site detail generation."""
        try:
            from backend.site_detail_processor import generate_site_detail, get_latest_available_day

            local_root = self.local_root_var.get().strip()
            latest_folder = get_latest_available_day(local_root, log_callback=self._log)

            if not latest_folder:
                self._log("❌ No data available", 'error')
                self.detail_status.set("Error - No data available")
                return

            self._log(f"📁 Using latest day: {os.path.basename(os.path.dirname(latest_folder))}")

            df = generate_site_detail(latest_folder, log_callback=self._log)

            if df is None or df.empty:
                self._log("❌ No site detail data generated", 'error')
                self.detail_status.set("Error - No data")
                return

            self.detail_df = df
            self._log(f"✅ Generated {len(df)} site details", 'success')

            self.root.after(0, self._update_detail_tree, df)

            # Save to CSV using CSVHistoryManager
            self.history_mgr.update_site_detail(df)
            self._log("✅ Site detail saved to CSV", 'success')

            self.detail_status.set(f"Ready - {len(df)} sites loaded and saved")

        except Exception as e:
            self._log_to_console(f"❌ Error generating site detail: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error: {e}", 'error')
            self.detail_status.set("Error")

    def _update_detail_tree(self, df):
        """Update the detail treeview with data."""
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)

        for idx, row in df.iterrows():
            values = [row.get(col, '') for col in SITE_DETAIL_HEADER]
            self.detail_tree.insert('', 'end', values=values)

    def _save_detail_to_excel(self):
        """Save the detail table to CSV/Excel."""
        if self.detail_df is None or self.detail_df.empty:
            messagebox.showwarning("No Data", "Please generate site detail first.")
            return

        try:
            self.history_mgr.update_site_detail(self.detail_df)
            excel_path = self.history_mgr.export_to_excel()
            self._log("✅ Site detail saved to CSV and Excel", 'success')
            messagebox.showinfo("Saved", f"Site detail saved to:\n{excel_path if excel_path else 'CSV files'}")
        except Exception as e:
            self._log_to_console(f"❌ Error saving: {e}")
            import traceback
            self._log_to_console(traceback.format_exc())
            self._log(f"❌ Error saving: {e}", 'error')
            messagebox.showerror("Error", f"Failed to save: {e}")

    def _copy_detail(self):
        """Copy detail table to clipboard as text."""
        if self.detail_df is None or self.detail_df.empty:
            messagebox.showwarning("No Data", "Please generate site detail first.")
            return

        text = "Site Detail Report\n"
        text += "=" * 80 + "\n"

        text += "\t".join(SITE_DETAIL_HEADER) + "\n"
        text += "-" * 80 + "\n"

        for _, row in self.detail_df.iterrows():
            values = [str(row.get(col, '')) for col in SITE_DETAIL_HEADER]
            text += "\t".join(values) + "\n"

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.detail_status.set("Copied to clipboard!")

    def _export_detail_csv(self):
        """Export detail table to CSV."""
        if self.detail_df is None or self.detail_df.empty:
            messagebox.showwarning("No Data", "Please generate site detail first.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save Site Detail as CSV"
        )

        if file_path:
            try:
                self.detail_df.to_csv(file_path, index=False)
                self._log(f"✅ Exported to {file_path}", 'success')
                self.detail_status.set(f"Exported to {os.path.basename(file_path)}")
                messagebox.showinfo("Exported", f"Data exported to {file_path}")
            except Exception as e:
                self._log(f"❌ Error exporting: {e}", 'error')
                messagebox.showerror("Error", f"Failed to export: {e}")

    # ---------- Log Methods ----------
    def _log(self, message, level='info'):
        """Add message to log queue for GUI update."""
        self.log_queue.put((message, level))

    def _poll_log_queue(self):
        """Periodically check log queue and update GUI."""
        try:
            while True:
                msg, level = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, msg + "\n", level)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_log_queue)

    def _copy_log(self):
        """Copy the log to clipboard."""
        self._log_to_console("📋 Copy Log clicked")
        self.root.clipboard_clear()
        self.root.clipboard_append(self.log_text.get(1.0, tk.END))
        self.status_var.set("Log copied to clipboard!")

    def _clear_log(self):
        """Clear the log."""
        self._log_to_console("🗑 Clear Log clicked")
        self.log_text.delete(1.0, tk.END)
        self.status_var.set("Log cleared")

    def _run_thread(self, target, *args):
        """Run a function in a separate thread."""

        def wrapper():
            try:
                self._log_to_console(f"🔄 Thread started: {target.__name__}")
                target(*args)
                self._log_to_console(f"✅ Thread completed: {target.__name__}")
            except Exception as e:
                self._log_to_console(f"❌ Thread error in {target.__name__}: {e}")
                import traceback
                self._log_to_console(traceback.format_exc())
                self._log(f"Unexpected error: {e}", 'error')
                self.status_var.set("Error")

        threading.Thread(target=wrapper, daemon=True).start()

    def _on_close(self):
        """Save config and close."""
        self._log_to_console("🚪 Application closing...")
        self.config.save()
        self._log_to_console("✅ Config saved")
        self.root.destroy()


# ---------------------------- Main Entry ----------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("🚀 LIBYANA NPM - MAIN APPLICATION")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    root = tk.Tk()
    app = LibyanaNPMApp(root)
    root.mainloop()