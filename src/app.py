"""
DataLens for Cyber Cell - Web Application.

A Streamlit-based web application for law enforcement cybercrime departments
to analyze and consolidate transaction data from Excel files.

Last Updated: 2026-04-27 - Account report now uses distinct Account Numbers
"""

import sys
import os

# Ensure the project root is on sys.path so `import src.xyz` works when
# Streamlit runs `src/app.py` as a script. This makes imports stable whether
# you run `python -m streamlit run src/app.py` from the repo root or from
# another working directory.
_THIS_FILE = os.path.abspath(__file__)
_PROJECT_ROOT = os.path.dirname(os.path.dirname(_THIS_FILE))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from typing import Optional, List, Dict, Any

from src.upload_service import UploadService
from src.column_detector import ColumnDetector
from src.data_processor import DataProcessor
from src.validation_engine import ValidationEngine
from src.aggregation_engine import AggregationEngine
from src.report_generator import ReportGenerator
from src.dashboard import Dashboard
from src.session_manager import SessionManager
from src.models import ColumnMapping, AggregatedAccount, ProcessingStats, ValidationResult
from src.district_data import render_district_download_page
from src.merge_files import render_merge_files_page
from src.excel_merger import render_excel_merger_page
from src.call_notice_data_merge import render_call_notice_merge_page
from src.transaction_matcher import render_transaction_matcher_page
from src.disputed_amount_matcher import render_disputed_amount_matcher_page
from src.money_transfer_dispute import render_money_transfer_dispute_page
from src.ack_bank_consolidator import render_ack_bank_consolidator_page
from src.database_service import DatabaseService
from src.report_service import ReportService
from src.districtwise import render_districtwise_page
from src.non_gujarat_filter import render_non_gujarat_filter_page
from src.column_selector import render_column_selector_page
from src.csv_fixer import render_csv_fixer_page
from src.drop_call_finder import render_drop_call_finder_page
from src.mo_finder import render_mo_finder_page
from src.amount_matcher import render_amount_matcher_page
from src.bank_ack_pivot import render_bank_ack_pivot_page
from src.ack_list_pivot import render_ack_list_pivot_page
from src.filter_by_entry_count import render_filter_by_entry_count_page
from src.filter_by_unique_ack import render_filter_by_unique_ack_page
from src.bulk_mysql_import import render_bulk_mysql_import_page
from src.mysql_database_viewer import render_mysql_database_viewer_page
from src.ai_sql_assistant import render_ai_sql_assistant_page
from src.automated_workflow import render_automated_workflow_page
from src.gujarat_account_formatter import render_gujarat_account_formatter_page
from src.column_transfer import render_column_transfer_page
from src.ui_styling import apply_custom_css, render_page_header_with_info
from src.smart_district_split import render_smart_district_split_page
from src.ifsc_pincode_district_split import render_ifsc_pincode_district_split_page
from src.daily_report_district_split import render_daily_report_district_split_page
from src.distinct_account_pivot import render_distinct_account_pivot_page
from src.top_10_suspect_accounts import render_top_10_suspect_accounts_page
from src.outsource_attendance import (
    render_attendance_admin_page,
    render_attendance_observer_page,
    render_outsource_login_page,
)

# Page configuration
st.set_page_config(
    page_title="DataLens for Cyber Cell",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom styling
apply_custom_css()

# Initialize services (cached to avoid recreation)
@st.cache_resource
def get_services():
    """Initialize and cache service instances."""
    return {
        'upload_service': UploadService(),
        'column_detector': ColumnDetector(),
        'data_processor': DataProcessor(),
        'validation_engine': ValidationEngine(),
        'aggregation_engine': AggregationEngine(),
        'report_generator': ReportGenerator(),
        'dashboard': Dashboard(),
        'session_manager': SessionManager()
    }


def init_session_state():
    """Initialize session state variables."""
    if 'current_page' not in st.session_state:
        st.session_state.current_page = 'upload'
    if 'uploaded_df' not in st.session_state:
        st.session_state.uploaded_df = None
    if 'filename' not in st.session_state:
        st.session_state.filename = None
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = None
    if 'cleaned_df' not in st.session_state:
        st.session_state.cleaned_df = None
    if 'validation_result' not in st.session_state:
        st.session_state.validation_result = None
    if 'aggregated_accounts' not in st.session_state:
        st.session_state.aggregated_accounts = None
    if 'processing_stats' not in st.session_state:
        st.session_state.processing_stats = None
    if 'processing_logs' not in st.session_state:
        st.session_state.processing_logs = []


def render_sidebar():
    """Render the navigation sidebar."""
    with st.sidebar:
        st.markdown(
            '<div class="sidebar-logo" aria-label="DataLens Cyber Cell Intelligence">'
            '<div class="sidebar-brand">'
            '<span class="brand-data">Data</span><span class="brand-lens">Lens</span>'
            '</div>'
            '<div class="brand-accent-line"><span></span><span></span><span></span><span></span></div>'
            '<div class="sidebar-subtitle">Cyber Cell Intelligence</div>'
            '</div>',
            unsafe_allow_html=True
        )
        st.markdown("---")
        
        # Page navigation
        pages = {
            'upload': '� Aggregate by Account',
            'attendance_admin': 'Attendance Admin',
            'attendance_observer': 'Observer Approvals',
            'outsource_login': 'Outsource Login',
            'district_download': '📍 Victim-Suspect Mapping & Filter by State/District',
            'top_10_suspect': '🎯 Top 20 Suspect Accounts from Layer 1',
            'districtwise': '📊 Split Data by Column',
            'smart_district_split': '🗺️ Smart District Split',
            'ifsc_pincode_split': '🏦 Split Excel Districtwise by IFSC/PIN Code',
            'filter_by_entry_count': '🔢 Filter by Entry Count',
            'filter_by_unique_ack': '🏦 Filter by Entry Count (Unique ACK Only)',
            'non_gujarat_filter': '🗺️ Non-Gujarat Filter',
            'amount_matcher': '💰 Add Disputed Amount to Pending/Unattended ZIP File',
            'bank_ack_pivot': '🏦 Bank ACK Pivot',
            'ack_list_pivot': '📋 ACK List Pivot',
            'report_generator': '📊 Account & Hold Amount Report Generator',
            'automated_workflow': '🔄 Automated Workflow',
            'gujarat_account_formatter': 'Gujarat Unique Account Output',
            'column_transfer': 'Add Columns by Match',
            'column_selector': '📋 Filter Excel File with the Columns You Need',
            'csv_fixer': 'CSV Fixer',
            'excel_merger': '📎 Merge Excel Files',
            'drop_call_finder': 'Drop Call Finder',
            'mo_finder': 'MO Finder',
            'call_notice_merge': '📞 Call Notice Data Mapping for Time Difference Between Call and Notice',
            'transaction_matcher': '🔄 Bring Disputed Amount to Unattended File (From Status-Wise Report)',
            'disputed_amount_matcher': '💰 Bring Disputed Amount to Any File with ACK No., Account Number & Transaction Amount',
            'money_transfer_dispute': '💸 Bring Disputed Amount to Money Transfer File',
            'ack_bank_consolidator': '📊 Acknowledgement Number & Bank-Wise Amount Consolidator',
            'bulk_mysql_import': '📊 Bulk MySQL Import',
            'mysql_database_viewer': '🗄️ MySQL Database Viewer',
            'ai_sql_assistant': '🤖 AI SQL Assistant',
            'distinct_account_pivot': '📊 Distinct Account Pivot',
            'view_database': '🗄️ View Database'
        }
        pages = dict(
            list(pages.items())[:6]
            + [('daily_report_district_split', 'District Wise Suspect Data Gujarat')]
            + list(pages.items())[6:]
        )

        clean_page_names = {
            'upload': 'Aggregate by Account',
            'attendance_admin': 'Attendance Admin',
            'attendance_observer': 'Observer Approvals',
            'outsource_login': 'Outsource Login',
            'district_download': 'Victim-Suspect Mapping',
            'top_10_suspect': 'Top 20 Suspect Accounts',
            'districtwise': 'Split Data by Column',
            'smart_district_split': 'Smart District Split',
            'ifsc_pincode_split': 'IFSC/PIN District Split',
            'filter_by_entry_count': 'Filter by Entry Count',
            'filter_by_unique_ack': 'Unique ACK Filter',
            'non_gujarat_filter': 'Non-Gujarat Filter',
            'amount_matcher': 'Disputed Amount Matcher',
            'bank_ack_pivot': 'Bank ACK Pivot',
            'ack_list_pivot': 'ACK List Pivot',
            'report_generator': 'Account & Hold Report',
            'automated_workflow': 'Automated Workflow',
            'gujarat_account_formatter': 'Gujarat Unique Account Output',
            'column_transfer': 'Add Columns by Match',
            'column_selector': 'Column Selector',
            'csv_fixer': 'CSV Fixer',
            'excel_merger': 'Merge Excel Files',
            'drop_call_finder': 'Drop Call Finder',
            'mo_finder': 'MO Finder',
            'call_notice_merge': 'Call Notice Mapping',
            'transaction_matcher': 'Transaction Matcher',
            'disputed_amount_matcher': 'ACK Account Matcher',
            'money_transfer_dispute': 'Money Transfer Dispute',
            'ack_bank_consolidator': 'ACK Bank Consolidator',
            'bulk_mysql_import': 'Bulk MySQL Import',
            'mysql_database_viewer': 'MySQL Database Viewer',
            'ai_sql_assistant': 'AI SQL Assistant',
            'distinct_account_pivot': 'Distinct Account Pivot',
            'view_database': 'View Database'
        }

        clean_page_names['daily_report_district_split'] = 'District Wise Suspect Data Gujarat'

        valid_page_keys = list(pages.keys())
        attendance_page_keys = ['attendance_admin', 'attendance_observer', 'outsource_login']
        top_page_keys = [
            'top_10_suspect',
            'automated_workflow',
            'gujarat_account_formatter',
            'column_transfer',
            'daily_report_district_split',
            'csv_fixer',
            'report_generator',
            'call_notice_merge',
            'drop_call_finder',
        ]

        def page_label(page_key):
            return clean_page_names.get(page_key, pages.get(page_key, page_key))

        def sidebar_section_title(title, detail=None):
            detail_html = f'<span>{detail}</span>' if detail else ''
            st.markdown(
                f'<div class="sidebar-section-title">{title}{detail_html}</div>',
                unsafe_allow_html=True,
            )

        st.markdown(
            """
            <style>
            [class*="st-key-attendance_nav_"] button {
                background: linear-gradient(135deg, #15251D, #11161E) !important;
                border: 1px solid #21C16B !important;
                border-left: 4px solid #FF9F0A !important;
                color: #F8F3EA !important;
                font-weight: 800 !important;
            }
            [class*="st-key-attendance_nav_"] button:hover {
                background: #1E3328 !important;
                border-color: #FF9F0A !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
        sidebar_section_title("Attendance Desk", "Protected")
        attendance_labels = {
            'attendance_admin': 'Admin Panel',
            'attendance_observer': 'Observer Desk',
            'outsource_login': 'Outsource Login',
        }
        for page_key in attendance_page_keys:
            if st.button(attendance_labels[page_key], key=f"attendance_nav_{page_key}", use_container_width=True):
                st.session_state.current_page = page_key
                st.rerun()
        st.caption("Password-protected attendance workflow.")
        st.markdown("---")

        sidebar_section_title("All Pages")
        ordered_page_keys = top_page_keys + [
            page_key
            for page_key in pages.keys()
            if page_key not in attendance_page_keys and page_key not in top_page_keys
        ]
        for page_key in ordered_page_keys:
            if page_key in attendance_page_keys:
                continue
            page_name = page_label(page_key)
            if st.button(page_name, key=f"nav_{page_key}", use_container_width=True):
                st.session_state.current_page = page_key
                st.rerun()
        
        st.markdown("---")
        
        # Session info
        st.caption("Session Info")
        if st.session_state.filename:
            st.caption(st.session_state.filename)
        
        st.markdown("---")
        
        # Data handling reminder
        st.caption("Security")
        st.caption("All data processed in-memory only.")
        st.caption("No data stored on servers.")
        
        # Reset button
        st.markdown("---")
        if st.button("Start Over", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

# (rest of file unchanged)
