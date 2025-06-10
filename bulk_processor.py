import sys
import os
import json
import sqlite3
import tempfile
import re
import io
import threading
import time
import logging
import decimal
import gc
import psutil
import shutil
import subprocess
from datetime import datetime, date
from typing import Dict, List, Optional, Any, Union
import fitz  # PyMuPDF
import pypdf_table_extraction
import pandas as pd
import numpy as np
from license_manager import get_license_manager

# Import factory modules for code deduplication
from common_factories import (
    TemplateFactory, DatabaseOperationFactory, UIMessageFactory,
    ValidationFactory, get_database_factory
)
from ui_component_factory import UIComponentFactory, LayoutFactory
from simplified_extraction_engine import get_extraction_engine
from pdf_extraction_utils import (
    extract_table, extract_tables, extract_invoice_tables as utils_extract_invoice_tables,
    clean_dataframe, DEFAULT_EXTRACTION_PARAMS,
    convert_display_to_pdf_coords, convert_pdf_to_display_coords, get_scale_factors,
    clear_extraction_cache, clear_extraction_cache_for_pdf, get_extraction_cache_stats
)
from multi_method_extraction import extract_with_method, cleanup_extraction

# Import standardized coordinate system - NO backward compatibility
from standardized_coordinates import StandardRegion, DatabaseConverter
from coordinate_boundary_converters import DatabaseBoundaryConverter
from region_utils import validate_rect
from error_handler import log_error, log_info, log_warning, handle_exception, ErrorContext
from extraction_params_utils import (
    normalize_extraction_params, prepare_section_params, create_standardized_extraction_call,
    ExtractionParamsHandler
)
from region_label_utils import (
    RegionLabelHandler, create_region_label, standardize_dataframe_labels, get_display_label
)
from split_screen_invoice_processor import SplitScreenInvoiceProcessor, RegionType
# Import invoice2data utilities
import invoice2data_utils
# Import unified invoice processing utilities
import invoice_processing_utils

# Import unified extraction components
try:
    from extraction_adapters import extract_from_database_template, process_invoice2data_unified
    from common_extraction_engine import get_extraction_engine
    from region_adapters import create_region_adapter
    UNIFIED_EXTRACTION_AVAILABLE = True
except ImportError as e:
    print(f"[WARNING] Unified extraction not available: {e}")
    UNIFIED_EXTRACTION_AVAILABLE = False

# Import cache manager
try:
    from cache_manager import get_cache_manager, register_temp_directory, register_cleanup_callback
    CACHE_MANAGER_AVAILABLE = True
except ImportError as e:
    print(f"[WARNING] Cache manager not available: {e}")
    CACHE_MANAGER_AVAILABLE = False

# Import standardized metadata handler
try:
    from standardized_metadata_handler import create_standard_metadata, format_metadata_for_output
    STANDARDIZED_METADATA_AVAILABLE = True
except ImportError as e:
    print(f"[WARNING] Standardized metadata handler not available: {e}")
    STANDARDIZED_METADATA_AVAILABLE = False

# Custom JSON encoder to handle non-serializable objects
class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, date)):
            return obj.isoformat()
        elif isinstance(obj, decimal.Decimal):
            return float(obj)
        elif hasattr(obj, 'to_dict'):
            return obj.to_dict()
        elif pd.isna(obj):
            return None
        try:
            return super().default(obj)
        except TypeError:
            return str(obj)

# Import invoice2data directly
try:
    from invoice2data import extract_data
    from invoice2data.extract.loader import read_templates
    from invoice2data.extract.invoice_template import InvoiceTemplate
    INVOICE2DATA_AVAILABLE = True
    invoice2data_version = "Using invoice2data directly"
except ImportError:
    INVOICE2DATA_AVAILABLE = False
    invoice2data_version = "Not available"
    print("Warning: invoice2data not available. Some features may not work.")
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QTableWidget,
    QTableWidgetItem,
    QProgressBar,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QStackedWidget,
    QFrame,
    QHeaderView,
    QGroupBox,
    QSplitter,
    QGridLayout,
    QSizePolicy,
    QListView,
    QStyle,
    QProxyStyle,
)
from PySide6.QtCore import Qt, Signal, QObject, QRect, QTimer
from PySide6.QtGui import QColor, QFont, QIcon
import pandas as pd
import time



class NoFrameStyle(QProxyStyle):
    def styleHint(self, hint, option=None, widget=None, returnData=None):
        if hint == QStyle.SH_ComboBox_Popup:
            return 0
        return super().styleHint(hint, option, widget, returnData)


class BulkProcessor(QWidget):
    # Define signals
    back_requested = Signal()  # Signal for navigating back to main dashboard
    go_back = Signal()  # Signal for navigating back to main dashboard

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.pdf_files = []
        self.processed_data = {}

        # Initialize stop flag for processing
        self.should_stop = False
        self.start_time = None

        # Memory management settings
        self.memory_threshold = 80  # Percentage of system memory that triggers cleanup
        self.batch_size = 10  # Number of PDFs to process in a batch
        self.current_memory_usage = 0  # Current memory usage in percentage
        self.memory_monitor_timer = None  # Timer for monitoring memory usage

        # Temporary directory for invoice files
        self.temp_dir = None
        self.invoice_txt_files = {}  # Map of PDF path to txt file path
        self.template_file_path = None

        # Initialize license manager
        self.license_manager = get_license_manager()
        self.license_info = {}
        self.file_limit = 0
        self.files_processed = 0
        self.license_valid = False

        # Define AI theme colors
        self.theme = {
            "primary": "#6366F1",       # Indigo
            "primary_dark": "#4F46E5",  # Darker indigo
            "secondary": "#10B981",     # Emerald
            "tertiary": "#8B5CF6",      # Violet
            "danger": "#EF4444",        # Red
            "warning": "#F59E0B",       # Amber
            "light": "#F9FAFB",         # Light gray
            "dark": "#111827",          # Dark gray
            "bg": "#F3F4F6",            # Background light gray
            "text": "#1F2937",          # Text dark
            "border": "#E5E7EB",        # Border light gray
        }

        # Set widget background
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {self.theme['bg']};
                color: {self.theme['text']};
                font-family: 'Segoe UI', Arial, sans-serif;
            }}

            QLabel {{
                font-size: 14px;
            }}

            QComboBox {{
                border: 1px solid {self.theme['border']};
                border-radius: 6px;
                padding: 8px 12px;
                background-color: white;
                min-height: 22px;
                selection-background-color: {self.theme['primary']};
            }}

            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 0px;
                border-top-right-radius: 6px;
                border-bottom-right-radius: 6px;
            }}

            QListWidget {{
                border: 1px solid {self.theme['border']};
                border-radius: 6px;
                background-color: white;
                padding: 8px;
                selection-background-color: {self.theme['primary']};
                selection-color: white;
            }}

            QTableWidget {{
                border: 1px solid {self.theme['border']};
                border-radius: 6px;
                background-color: white;
                gridline-color: {self.theme['border']};
                selection-background-color: {self.theme['primary']};
                selection-color: white;
            }}

            QTableWidget::item {{
                padding: 6px;
            }}

            QHeaderView::section {{
                background-color: {self.theme['light']};
                border: none;
                padding: 8px;
                font-weight: bold;
                color: {self.theme['text']};
                border-right: 1px solid {self.theme['border']};
                border-bottom: 1px solid {self.theme['border']};
            }}

            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: {self.theme['light']};
                height: 12px;
                text-align: center;
            }}

            QProgressBar::chunk {{
                background-color: {self.theme['primary']};
                border-radius: 4px;
            }}
        """)

        self.init_ui()
        self.load_license_info()  # Load license information after UI is initialized
        self.load_templates()  # Load templates when initializing

        # Register with cache manager for cleanup
        if CACHE_MANAGER_AVAILABLE:
            try:
                register_cleanup_callback(self.cleanup_bulk_processor_caches)
                print("[DEBUG] Registered BulkProcessor cleanup callback with cache manager")
            except Exception as e:
                print(f"[WARNING] Failed to register with cache manager: {e}")

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)  # Reduced spacing
        layout.setContentsMargins(8, 2, 8, 8)  # Reduced top margin

        # Header section - Title, navigation buttons, and license info
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)  # No margins at all
        header_layout.setSpacing(8)  # Minimal spacing

        # Back button on the left with white background and black text
        back_btn = QPushButton("← Back", self)
        back_btn.clicked.connect(self.navigate_back)
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: black;
                padding: 4px 12px;
                border-radius: 4px;
                border: none;
                font-weight: bold;
                min-height: 28px;
                outline: none;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
            }
            QPushButton:pressed {
                background-color: #e0e0e0;
                padding-top: 5px;
                padding-left: 13px;
            }
            QPushButton:focus {
                outline: none;
                border: none;
            }
        """)
        header_layout.addWidget(back_btn)

        # Reset screen button next to back button
        reset_btn = QPushButton("Reset Screen", self)
        reset_btn.clicked.connect(self.reset_screen)
        reset_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['danger']};
                color: white;
                padding: 4px 12px;
                border-radius: 4px;
                border: none;
                font-weight: bold;
                min-height: 28px;
                outline: none;
            }}
            QPushButton:hover {{
                background-color: {self.theme['danger'] + 'E0'};
            }}
            QPushButton:pressed {{
                background-color: {self.theme['danger'] + 'C0'};
                padding-top: 5px;
                padding-left: 13px;
            }}
            QPushButton:focus {{
                outline: none;
                border: none;
            }}
        """)
        header_layout.addWidget(reset_btn)

        # Add stretch to push title to center
        header_layout.addStretch(1)

        # Title in the center with improved styling
        title_label = QLabel("Bulk Extractor", self)
        title_label.setStyleSheet(f"""
            font-size: 24px;
            font-weight: bold;
            color: white;
            background-color: {self.theme['primary']};
            margin: 0;
            padding: 4px 12px;
            border-radius: 6px;
            text-align: center;
        """)
        title_label.setFixedHeight(36)  # Better height for visibility
        title_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(title_label)

        # Add stretch to balance the layout
        header_layout.addStretch(1)

        # License info on the right (minimal height)
        license_info_frame = QFrame()
        license_info_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {self.theme['light']};
                border-radius: 2px;  /* Minimal radius */
                padding: 1px 4px;  /* Minimal padding */
            }}
        """)
        license_info_frame.setFixedHeight(40)  # Force minimal height
        license_info_layout = QHBoxLayout(license_info_frame)
        license_info_layout.setSpacing(4)  # Minimal spacing
        license_info_layout.setContentsMargins(2, 2, 2, 2)  # Minimal margins

        # License status (tiny font)
        license_label = QLabel("License:", self)
        license_label.setStyleSheet("font-weight: bold; font-size: 10px;")
        license_label.setFixedHeight(20)
        self.license_status_label = QLabel("Loading...", self)
        self.license_status_label.setStyleSheet(f"color: {self.theme['tertiary']}; font-weight: bold; font-size: 10px;")
        self.license_status_label.setFixedHeight(20)

        # Usage info (tiny font)
        usage_label = QLabel("Files:", self)
        usage_label.setStyleSheet("font-weight: bold; font-size: 10px;")
        usage_label.setFixedHeight(10)
        self.license_usage_label = QLabel("0/0", self)
        self.license_usage_label.setStyleSheet(f"color: {self.theme['tertiary']}; font-weight: bold; font-size: 10px;")
        self.license_usage_label.setFixedHeight(10)

        # Expiry info (tiny font)
        expiry_label = QLabel("Valid Until:", self)
        expiry_label.setStyleSheet("font-weight: bold; font-size: 10px;")
        expiry_label.setFixedHeight(10)
        self.license_expiry_label = QLabel("Unknown", self)
        self.license_expiry_label.setStyleSheet(f"color: {self.theme['tertiary']}; font-weight: bold; font-size: 10px;")
        self.license_expiry_label.setFixedHeight(10)

        # Add to license info layout
        license_info_layout.addWidget(license_label)
        license_info_layout.addWidget(self.license_status_label)
        license_info_layout.addWidget(usage_label)
        license_info_layout.addWidget(self.license_usage_label)
        license_info_layout.addWidget(expiry_label)
        license_info_layout.addWidget(self.license_expiry_label)

        header_layout.addWidget(license_info_frame)
        layout.addLayout(header_layout)

        # Create a horizontal splitter to divide the screen left/right
        main_splitter = QSplitter(Qt.Horizontal)
        main_splitter.setHandleWidth(1)
        main_splitter.setStyleSheet(f"""
            QSplitter::handle {{
                background-color: {self.theme['border']};
            }}
        """)

        # LEFT SECTION - Extraction controls
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(16)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Template selection with card style
        template_card = QFrame(self)
        template_card.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border: 1px solid {self.theme['border']};
                padding: 16px;
            }}
        """)
        template_layout = QVBoxLayout(template_card)
        template_label = QLabel("Select Template", self)
        template_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        template_layout.addWidget(template_label)

        template_input_layout = QHBoxLayout()
        self.template_combo = QComboBox(self)
        self.template_combo.setMinimumHeight(36)

        # Apply custom style to remove frame
        self.template_combo.setStyle(NoFrameStyle())

        # Create and set a custom list view for the combo box
        list_view = QListView()
        list_view.setFrameShape(QListView.NoFrame)
        self.template_combo.setView(list_view)

        # Styled dropdown with proper border
        self.template_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #E5E7EB;
                border-radius: 6px;
                padding: 8px 12px;
                background-color: white;
                min-height: 36px;
                color: #1F2937;
                font-size: 14px;
            }

            QComboBox:hover, QComboBox:focus {
                border: 1px solid #6366F1;
            }

            QComboBox::drop-down {
                border: none;
                width: 20px;
            }

            QComboBox::down-arrow {
                image: none;
            }

            QComboBox QAbstractItemView {
                border: 1px solid #E5E7EB;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
                outline: none;
            }

            QComboBox QAbstractItemView::item {
                border-bottom: 1px solid #F3F4F6;
                padding: 8px 12px;
                min-height: 30px;
                color: #1F2937;
            }

            QComboBox QAbstractItemView::item:last-child {
                border-bottom: none;
            }

            QComboBox QAbstractItemView::item:hover {
                background-color: #F3F4F6;
            }

            QComboBox QAbstractItemView::item:selected {
                background-color: #EEF2FF;
                color: #4F46E5;
            }
        """)

        # Add refresh button
        refresh_btn = QPushButton("Refresh", self)
        refresh_btn.clicked.connect(self.load_templates)
        refresh_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['primary']};
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                min-height: 36px;
                min-width: 100px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['primary_dark']};
            }}
            QPushButton:pressed {{
                background-color: {self.theme['primary_dark']};
                padding-top: 9px;
                padding-left: 17px;
            }}
        """)

        template_input_layout.addWidget(self.template_combo, 1)
        template_input_layout.addWidget(refresh_btn, 0)
        template_layout.addLayout(template_input_layout)

        # Multi-page info label with modern style
        self.multi_page_label = QLabel("Multi-page support: Enabled ✓", self)
        self.multi_page_label.setStyleSheet(f"""
            color: {self.theme['secondary']};
            font-weight: bold;
            padding: 4px 8px;
            background-color: {self.theme['secondary'] + '20'};
                border-radius: 4px;
        """)
        template_layout.addWidget(self.multi_page_label)

        # Status and Progress
        status_layout = QHBoxLayout()
        # Status indicator without any background or border
        self.status_label = QLabel("Ready", self)
        self.status_label.setStyleSheet("color: #00B8A9;")
        status_layout.addWidget(self.status_label)

        # Processing time label - simple text only
        self.processing_time_label = QLabel("", self)
        self.processing_time_label.setStyleSheet("color: #1F2937;")
        status_layout.addWidget(self.processing_time_label)

        # Add stretch to push labels to the left
        status_layout.addStretch()
        template_layout.addLayout(status_layout)

        # Progress bar and stop button in a layout
        progress_layout = QHBoxLayout()

        # Progress bar with modern style
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setMaximumHeight(8)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: none;
                border-radius: 4px;
                background-color: {self.theme['light']};
                height: 8px;
            }}
            QProgressBar::chunk {{
                background-color: {self.theme['primary']};
                border-radius: 4px;
            }}
        """)
        progress_layout.addWidget(self.progress_bar, 1)

        # Stop button
        self.stop_button = QPushButton("Stop", self)
        self.stop_button.clicked.connect(self.stop_processing)
        self.stop_button.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['danger']};
                color: white;
                padding: 4px 12px;
                border-radius: 4px;
                font-weight: bold;
                min-height: 24px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['danger'] + 'CC'};
            }}
            QPushButton:pressed {{
                background-color: {self.theme['danger'] + 'AA'};
            }}
        """)
        self.stop_button.setVisible(False)
        progress_layout.addWidget(self.stop_button)

        template_layout.addLayout(progress_layout)

        left_layout.addWidget(template_card)

        # File Operations section
        file_card = QFrame(self)
        file_card.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border: 1px solid {self.theme['border']};
                padding: 16px;
            }}
        """)
        file_layout = QVBoxLayout(file_card)

        file_title = QLabel("PDF Files", self)
        file_title.setStyleSheet("font-weight: bold; font-size: 16px;")
        file_layout.addWidget(file_title)

        # File list with better spacing
        self.file_list = QListWidget(self)
        self.file_list.setMinimumHeight(200)
        self.file_list.setStyleSheet(f"""
            QListWidget {{
                border: 1px solid {self.theme['border']};
                border-radius: 6px;
                background-color: white;
                padding: 4px;
            }}
            QListWidget::item {{
                border-bottom: 1px solid {self.theme['border'] + '50'};
                padding: 6px;
            }}
            QListWidget::item:selected {{
                background-color: {self.theme['primary'] + '30'};
                color: {self.theme['text']};
                border-radius: 4px;
            }}
        """)
        file_layout.addWidget(self.file_list)

        # Buttons for file operations
        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)

        add_files_btn = QPushButton("Add Files", self)
        add_files_btn.clicked.connect(self.add_files)
        add_files_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['primary']};
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                min-height: 36px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['primary_dark']};
            }}
            QPushButton:pressed {{
                background-color: {self.theme['primary_dark']};
                padding-top: 9px;
                padding-left: 17px;
            }}
        """)

        clear_files_btn = QPushButton("Clear Files", self)
        clear_files_btn.clicked.connect(self.clear_files)
        clear_files_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: white;
                color: {self.theme['text']};
                padding: 8px 16px;
                border-radius: 6px;
                border: 1px solid {self.theme['border']};
                font-weight: bold;
                min-height: 36px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['light']};
            }}
            QPushButton:pressed {{
                background-color: {self.theme['border']};
                padding-top: 9px;
                padding-left: 17px;
            }}
        """)

        process_btn = QPushButton("Process Files", self)
        process_btn.clicked.connect(self.process_files)
        process_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['primary']};
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                min-height: 36px;
            }}
            QPushButton:hover {{
                background-color: {self.theme['primary_dark']};
            }}
            QPushButton:pressed {{
                background-color: {self.theme['primary_dark']};
                padding-top: 9px;
                padding-left: 17px;
            }}
        """)

        button_layout.addWidget(add_files_btn)
        button_layout.addWidget(clear_files_btn)
        button_layout.addWidget(process_btn)
        file_layout.addLayout(button_layout)
        left_layout.addWidget(file_card)

        # Navigation section removed (moved to top)

        # Add left widget to splitter
        main_splitter.addWidget(left_widget)

        # RIGHT SECTION - Extraction results
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(16)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Results section
        results_card = QFrame(self)
        results_card.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border: 1px solid {self.theme['border']};
                padding: 16px;
            }}
        """)
        results_layout = QVBoxLayout(results_card)

        results_title = QLabel("Extraction Results", self)
        results_title.setStyleSheet("font-weight: bold; font-size: 16px;")
        results_layout.addWidget(results_title)

        # Summary statistics panel
        summary_frame = QFrame()
        summary_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {self.theme['light']};
                border-radius: 6px;
                padding: 12px;
                margin-bottom: 12px;
            }}
            QLabel {{
                font-size: 13px;
            }}
        """)
        summary_layout = QGridLayout(summary_frame)
        summary_layout.setSpacing(12)

        # Add summary statistics labels
        processed_label = QLabel("Processed Files:", self)
        processed_label.setStyleSheet("font-weight: bold;")
        self.processed_count = QLabel("0", self)
        self.processed_count.setStyleSheet(f"color: {self.theme['primary']}; font-weight: bold;")

        success_label = QLabel("Successful:", self)
        success_label.setStyleSheet("font-weight: bold;")
        self.success_count = QLabel("0", self)
        self.success_count.setStyleSheet(f"color: {self.theme['secondary']}; font-weight: bold;")

        failed_label = QLabel("Failed:", self)
        failed_label.setStyleSheet("font-weight: bold;")
        self.failed_count = QLabel("0", self)
        self.failed_count.setStyleSheet(f"color: {self.theme['danger']}; font-weight: bold;")

        total_rows_label = QLabel("Total Rows Extracted:", self)
        total_rows_label.setStyleSheet("font-weight: bold;")
        self.total_rows_count = QLabel("0", self)
        self.total_rows_count.setStyleSheet("font-weight: bold;")

        # Memory usage and cache size labels moved to bottom of screen

        # License information moved to top right of the screen

        # Add labels to grid layout in a single row
        summary_layout.addWidget(processed_label, 0, 0)
        summary_layout.addWidget(self.processed_count, 0, 1)
        summary_layout.addWidget(success_label, 0, 2)
        summary_layout.addWidget(self.success_count, 0, 3)
        summary_layout.addWidget(failed_label, 0, 4)
        summary_layout.addWidget(self.failed_count, 0, 5)
        summary_layout.addWidget(total_rows_label, 0, 6)
        summary_layout.addWidget(self.total_rows_count, 0, 7)

        # License information grid layout removed (moved to top right)

        results_layout.addWidget(summary_frame)

        # Results table with modern style
        self.results_table = QTableWidget(self)
        self.results_table.setColumnCount(7)  # Added one more column for validation status
        self.results_table.setHorizontalHeaderLabels(
            [
                "File Name",
                "Extraction Status",
                "Validation Status",  # New column
                "PDF Pages",
                "Header Data Rows",
                "Line Items Rows",
                "Summary Data Rows",
            ]
        )

        # Set up horizontal header with better styling
        header = self.results_table.horizontalHeader()

        # Configure header behavior
        header.setVisible(True)
        header.setHighlightSections(False)
        header.setStretchLastSection(False)
        header.setSectionsMovable(False)

        # Set column stretch factors and resize modes
        # Define proportional stretch factors for each column
        stretch_factors = [3, 2, 2, 1, 1.5, 1.5, 1.5]  # Proportional factors

        # Set minimum widths to ensure columns don't get too small
        min_widths = [150, 100, 100, 60, 80, 80, 80]

        # Apply stretch mode to all columns
        for i, (factor, min_width) in enumerate(zip(stretch_factors, min_widths)):
            self.results_table.setColumnWidth(i, min_width)  # Set initial width
            header.setSectionResizeMode(i, QHeaderView.Stretch)  # Use stretch mode

        # Set the stretch factor for each column
        total_stretch = sum(stretch_factors)
        for i, factor in enumerate(stretch_factors):
            header.resizeSection(i, int((factor / total_stretch) * 1000))  # Use 1000 as base width

        # Update table stylesheet with more prominent header styling and proper alignment
        self.results_table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {self.theme['border']};
                border-radius: 6px;
                background-color: white;
                gridline-color: {self.theme['border']};
            }}

            QHeaderView::section {{
                background-color: white;
                color: {self.theme['text']};
                font-weight: bold;
                font-size: 13px;
                padding: 8px;
                border: 1px solid {self.theme['border']};
                min-height: 30px;
                max-height: 30px;
            }}

            QHeaderView::section:horizontal {{
                border-top: 1px solid {self.theme['border']};
                text-align: left;
                padding-left: 12px;
            }}

            QTableWidget::item {{
                padding: 8px;
                border-bottom: 1px solid {self.theme['border']};
                text-align: left;
                padding-left: 12px;
            }}

            QTableWidget::item:selected {{
                background-color: {self.theme['primary'] + '15'};
                color: {self.theme['text']};
            }}
        """)

        # Additional header settings
        header.setMinimumHeight(40)
        header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        # Enable features for better usability
        self.results_table.setAlternatingRowColors(False)  # Disable alternating row colors
        self.results_table.setShowGrid(True)
        self.results_table.setGridStyle(Qt.SolidLine)
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setSelectionMode(QTableWidget.SingleSelection)
        self.results_table.setMinimumHeight(300)

        # Set table size policy to expand properly
        self.results_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Ensure the table is properly contained in a frame
        table_frame = QFrame()
        table_frame.setStyleSheet(f"""
            QFrame {{
                border: 1px solid {self.theme['border']};
                border-radius: 6px;
                background-color: white;
                padding: 1px;
            }}
        """)
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(1, 1, 1, 1)
        table_layout.addWidget(self.results_table)

        results_layout.addWidget(table_frame)

        # Memory and cache section - horizontal layout
        memory_cache_layout = QHBoxLayout()
        memory_cache_layout.setSpacing(16)
        memory_cache_layout.setContentsMargins(0, 8, 0, 0)  # Add a bit of top margin

        # Memory usage section
        memory_section = QHBoxLayout()
        memory_label = QLabel("Memory Usage:", self)
        memory_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.memory_usage_label = QLabel("0%", self)
        self.memory_usage_label.setStyleSheet(f"color: {self.theme['tertiary']}; font-weight: bold; font-size: 14px;")
        memory_section.addWidget(memory_label)
        memory_section.addWidget(self.memory_usage_label)
        memory_cache_layout.addLayout(memory_section)

        # Add some spacing
        memory_cache_layout.addSpacing(20)

        # Cache size section
        cache_section = QHBoxLayout()
        cache_label = QLabel("Cache Size:", self)
        cache_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.cache_size_label = QLabel("0 items", self)
        self.cache_size_label.setStyleSheet(f"color: {self.theme['tertiary']}; font-weight: bold; font-size: 14px;")
        cache_section.addWidget(cache_label)
        cache_section.addWidget(self.cache_size_label)
        memory_cache_layout.addLayout(cache_section)

        # Add stretch to push buttons to the right
        memory_cache_layout.addStretch(1)

        # Validate & Export button
        validate_btn = QPushButton("Validate and Export", self)
        validate_btn.clicked.connect(self.open_validation_screen)
        validate_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['tertiary']};
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                min-height: 36px;
                min-width: 180px;
            }}
            QPushButton:hover {{
                background-color: #7C3AED;
            }}
            QPushButton:pressed {{
                background-color: #6D28D9;
                padding-top: 9px;
                padding-left: 17px;
            }}
        """)
        memory_cache_layout.addWidget(validate_btn)

        results_layout.addLayout(memory_cache_layout)
        right_layout.addWidget(results_card)

        # Add right widget to splitter
        main_splitter.addWidget(right_widget)

        # Set the splitter proportions
        main_splitter.setSizes([500, 500])

        # Add splitter to main layout
        layout.addWidget(main_splitter)

        self.setLayout(layout)

    def load_license_info(self):
        """Load license information from the license manager"""
        try:
            # User ID is already set in get_license_manager function

            # Verify license
            is_valid, message = self.license_manager.verify_license()
            self.license_valid = is_valid

            # Get license info
            self.license_info = self.license_manager.get_license_info()

            # Get file limit
            self.file_limit = self.license_info.get("file_limit", 0)

            # Get files processed from the database
            self.files_processed = self.license_manager.get_files_processed()

            # Update UI
            if is_valid:
                self.license_status_label.setText("Valid")
                self.license_status_label.setStyleSheet(f"color: {self.theme['secondary']}; font-weight: bold;")
            else:
                self.license_status_label.setText("Invalid")
                self.license_status_label.setStyleSheet(f"color: {self.theme['danger']}; font-weight: bold;")

            # Update file limit display
            if self.file_limit <= 0:
                self.license_usage_label.setText("Unlimited")
            else:
                self.license_usage_label.setText(f"{self.files_processed}/{self.file_limit}")

            # Update expiry date
            if "expiry_date" in self.license_info:
                try:
                    expiry_date = datetime.fromisoformat(self.license_info["expiry_date"])
                    self.license_expiry_label.setText(expiry_date.strftime("%Y-%m-%d"))
                except (ValueError, TypeError):
                    self.license_expiry_label.setText("Unknown")
            else:
                self.license_expiry_label.setText("Never")

            print(f"License loaded: Valid={is_valid}, File limit={self.file_limit}, Files processed={self.files_processed}")

        except Exception as e:
            handle_exception(
                func_name="load_license_info",
                exception=e,
                context={"license_manager": str(type(self.license_manager))}
            )

            # Set default values
            self.license_valid = False
            self.file_limit = 0
            self.files_processed = 0
            self.license_status_label.setText("Error")
            self.license_status_label.setStyleSheet(f"color: {self.theme['danger']}; font-weight: bold;")
            self.license_usage_label.setText("0/0")
            self.license_expiry_label.setText("Unknown")

    def update_license_usage(self, additional_files):
        """Update the license usage counter"""
        # Update the database
        success = self.license_manager.update_files_processed(additional_files)

        if success:
            # Reload the current count from the database
            self.files_processed = self.license_manager.get_files_processed()

            # Update UI
            if self.file_limit <= 0:
                self.license_usage_label.setText("Unlimited")
            else:
                self.license_usage_label.setText(f"{self.files_processed}/{self.file_limit}")

            # Check if we're approaching the limit
            if 0 < self.file_limit < self.files_processed + 10:
                # We're getting close to the limit
                remaining = self.file_limit - self.files_processed
                if remaining > 0:
                    QMessageBox.warning(
                        self,
                        "License Limit Warning",
                        f"You are approaching your file processing limit. Only {remaining} files remaining."
                    )
        else:
            print("Failed to update license usage in the database")

    def check_license_limit(self, file_count):
        """Check if processing the given number of files is allowed by the license"""
        # Use the license manager to check the limit
        is_allowed, message = self.license_manager.check_bulk_limit(file_count)

        if not is_allowed:
            QMessageBox.critical(
                self,
                "License Limit Error",
                message
            )
            return False

        return True

    def load_templates(self):
        """Load templates from the database"""
        try:
            print("\nLoading templates for bulk processing...")
            conn = sqlite3.connect("invoice_templates.db")
            cursor = conn.cursor()

            # Check if the templates table exists
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='templates'"
            )
            if not cursor.fetchone():
                print("Templates table does not exist in the database")
                conn.close()
                QMessageBox.warning(
                    self,
                    "No Templates",
                    "No template table found in the database. Please create templates first.",
                )
                return

            # Get table columns to handle different database schemas
            cursor.execute("PRAGMA table_info(templates)")
            columns = cursor.fetchall()
            column_names = [col[1] for col in columns]
            print(f"Available template columns: {column_names}")

            # Build query dynamically based on available columns
            select_columns = ["id"]
            if "name" in column_names:
                select_columns.append("name")
            else:
                select_columns.append("'Unnamed'")

            if "template_type" in column_names:
                select_columns.append("template_type")
            else:
                select_columns.append("'single'")

            if "page_count" in column_names:
                select_columns.append("page_count")
            else:
                select_columns.append("1")

            # Add creation_date for sorting
            if "creation_date" in column_names:
                select_columns.append("creation_date")

            # Build query with ORDER BY creation_date DESC if available
            if "creation_date" in column_names:
                query = f"SELECT {', '.join(select_columns)} FROM templates ORDER BY creation_date DESC"
            else:
                query = f"SELECT {', '.join(select_columns)} FROM templates"
            print(f"Query: {query}")

            # Execute the query
            cursor.execute(query)
            templates = cursor.fetchall()
            print(f"Found {len(templates)} templates")

            # Print template details for debugging
            for template in templates:
                template_id = template[0]
                template_name = template[1] if len(template) > 1 else "Unnamed"
                template_type = template[2] if len(template) > 2 else "single"
                page_count = template[3] if len(template) > 3 else 1
                print(
                    f"  Template: {template_id}, {template_name}, {template_type}, {page_count} pages"
                )

            # Clear and reload the combo box
            self.template_combo.clear()

            # Add a placeholder item with no selection
            self.template_combo.addItem("-- Select a template --", None)

            if not templates:
                print("No templates found in database")
                self.template_combo.addItem("No templates available", None)
                self.multi_page_label.setText("Multi-page support: No templates found")
                self.multi_page_label.setStyleSheet("color: orange;")
            else:
                has_multi_page = False
                for template in templates:
                    template_id = template[0]
                    template_name = template[1] if len(template) > 1 else "Unnamed"
                    template_type = template[2] if len(template) > 2 else "single"
                    page_count = template[3] if len(template) > 3 else 1

                    if template_type == "multi":
                        has_multi_page = True
                        display_text = f"{template_name} ({template_type.title()}, {page_count} pages)"
                    else:
                        display_text = f"{template_name} ({template_type.title()})"

                    self.template_combo.addItem(display_text, template_id)
                    print(
                        f"Added template to dropdown: {display_text}, ID: {template_id}"
                    )

                # Update multi-page indicator
                if has_multi_page:
                    self.multi_page_label.setText("Multi-page support: Enabled ✓")
                    self.multi_page_label.setStyleSheet(
                        "color: green; font-weight: bold;"
                    )
                else:
                    self.multi_page_label.setText(
                        "Multi-page support: No multi-page templates found"
                    )
                    self.multi_page_label.setStyleSheet("color: orange;")

            conn.close()
            print("Finished loading templates")

        except sqlite3.Error as e:
            error_msg = f"Database error while loading templates: {str(e)}"
            print(error_msg)
            QMessageBox.critical(self, "Database Error", error_msg)
            self.multi_page_label.setText("Multi-page support: Database error")
            self.multi_page_label.setStyleSheet("color: red;")
        except Exception as e:
            error_msg = f"Failed to load templates: {str(e)}"
            print(error_msg)
            import traceback

            traceback.print_exc()
            QMessageBox.critical(self, "Error", error_msg)
            self.multi_page_label.setText("Multi-page support: Error loading templates")
            self.multi_page_label.setStyleSheet("color: red;")

    def process_files(self):
        """Process selected PDF files with the selected template"""
        # Validate files and template selection
        if not self.pdf_files:
            QMessageBox.warning(self, "Warning", "Please add PDF files first")
            return

        template_id = self.get_selected_template_id()
        if not template_id:
            QMessageBox.warning(self, "Warning", "Please select a template")
            return

        # Check license limit
        if not self.check_license_limit(len(self.pdf_files)):
            return

        # Reset counters and displays
        self.status_label.setText("Processing files...")
        self.results_table.setRowCount(0)
        self.processed_count.setText("0")
        self.success_count.setText("0")
        self.failed_count.setText("0")
        self.total_rows_count.setText("0")
        self.progress_bar.setMaximum(len(self.pdf_files))
        self.progress_bar.setValue(0)

        # Reset stop flag and show stop button
        self.should_stop = False
        self.stop_button.setVisible(True)

        # Start the timer
        self.start_time = time.time()
        self.processing_time_timer = QTimer(self)
        self.processing_time_timer.timeout.connect(self.update_processing_time)
        self.processing_time_timer.start(1000)  # Update every second

        # Start memory monitoring
        self.start_memory_monitoring()

        try:
            # Initialize counters for summary statistics
            processed_count = 0
            success_count = 0
            failed_count = 0
            total_rows = 0

            # Process PDFs in batches to manage memory usage
            for batch_start in range(0, len(self.pdf_files), self.batch_size):
                # Check if processing should stop
                if self.should_stop:
                    self.status_label.setText("Processing stopped by user")
                    break

                # Get the current batch of PDFs
                batch_end = min(batch_start + self.batch_size, len(self.pdf_files))
                current_batch = self.pdf_files[batch_start:batch_end]

                self.status_label.setText(f"Processing batch {batch_start//self.batch_size + 1}/{(len(self.pdf_files)-1)//self.batch_size + 1}...")

                # Process each PDF file in the current batch
                for index, pdf_path in enumerate(current_batch):
                    # Check if processing should stop
                    if self.should_stop:
                        self.status_label.setText("Processing stopped by user")
                        break

                    # Check memory usage and clean up if necessary
                    self.check_memory_usage()

                    print(
                        f"\nProcessing file {batch_start + index + 1}/{len(self.pdf_files)}: {pdf_path}"
                    )
                    self.status_label.setText(f"Processing: {os.path.basename(pdf_path)}")

                    try:
                        # Get actual PDF page count first
                        with fitz.open(pdf_path) as pdf:
                            actual_page_count = len(pdf)

                        # Extract tables from the PDF
                        results = self.extract_invoice_tables(pdf_path, template_id)

                        processed_count += 1
                        self.processed_count.setText(str(processed_count))

                        if results:
                            # Get template_type from the selected template name
                            template_display_text = self.template_combo.currentText()
                            template_type = "single"  # Default
                            if "Multi" in template_display_text:
                                template_type = "multi"

                            # Get the overall extraction status
                            extraction_status = results.get("extraction_status", {})
                            overall_status = extraction_status.get("overall", "failed")

                            # Store the results with correct page count and template type
                            self.processed_data[pdf_path] = {
                                "pdf_page_count": actual_page_count,  # Use actual page count from PDF
                                "template_type": template_type,  # Add template type
                                "header": results.get("header_tables", []),
                                "items": results.get("items_tables", []),
                                "summary": results.get("summary_tables", []),
                                "extraction_status": extraction_status  # Add extraction status
                            }

                            # Add to results table with correct counts
                            row = self.results_table.rowCount()
                            self.results_table.insertRow(row)

                            # File name
                            file_item = QTableWidgetItem(os.path.basename(pdf_path))
                            self.results_table.setItem(row, 0, file_item)

                            # Determine success status based on extraction_status
                            header_count = sum(
                                len(df)
                                for df in results.get("header_tables", [])
                                if df is not None and not df.empty
                            )
                            item_count = sum(
                                len(df)
                                for df in results.get("items_tables", [])
                                if df is not None and not df.empty
                            )
                            summary_count = sum(
                                len(df)
                                for df in results.get("summary_tables", [])
                                if df is not None and not df.empty
                            )

                            # Set status based on overall extraction status
                            if overall_status == "success":
                                status_text = "Success"
                                status_type = "success"
                                # Update success counter
                                success_count += 1
                                self.success_count.setText(str(success_count))
                            elif overall_status == "partial":
                                # Create a more detailed status message
                                partial_sections = []
                                if extraction_status.get("header") in ["success", "partial"]:
                                    partial_sections.append("Header")
                                if extraction_status.get("items") in ["success", "partial"]:
                                    partial_sections.append("Items")
                                if extraction_status.get("summary") in ["success", "partial"]:
                                    partial_sections.append("Summary")

                                status_text = f"Partial: {', '.join(partial_sections)}"
                                status_type = "partial"
                                # Update success counter but also track as partial
                                success_count += 1
                                self.success_count.setText(str(success_count))
                            else:
                                # Try to determine why it failed
                                if actual_page_count == 0:
                                    status_text = "Failed: Could not read PDF"
                                elif not any([header_count, item_count, summary_count]):
                                    status_text = "Failed: No data extracted"
                                else:
                                    status_text = "Failed: Extraction errors"

                                status_type = "failed"
                                # Update failed counter
                                failed_count += 1
                                self.failed_count.setText(str(failed_count))

                            status_item = QTableWidgetItem(status_text)
                            status_item.setData(Qt.UserRole, status_type)
                            self.results_table.setItem(row, 1, status_item)

                            # Validation Status - Set to "Pending" initially
                            validation_status_item = QTableWidgetItem("Pending")
                            validation_status_item.setData(Qt.UserRole, "pending")
                            validation_status_item.setForeground(QColor("#F59E0B"))  # Amber color for pending
                            self.results_table.setItem(row, 2, validation_status_item)

                            # PDF Pages - Use actual page count
                            self.results_table.setItem(
                                row, 3, QTableWidgetItem(str(actual_page_count))
                            )

                            # Header Rows - sum of rows in all header tables
                            self.results_table.setItem(
                                row, 4, QTableWidgetItem(str(header_count))
                            )

                            # Item Rows - sum of rows in all item tables
                            self.results_table.setItem(
                                row, 5, QTableWidgetItem(str(item_count))
                            )

                            # Summary Rows - sum of rows in all summary tables
                            self.results_table.setItem(
                                row, 6, QTableWidgetItem(str(summary_count))
                            )

                            # Update total rows counter
                            file_total_rows = header_count + item_count + summary_count
                            total_rows += file_total_rows
                            self.total_rows_count.setText(str(total_rows))
                        else:
                            # Add error to results table
                            row = self.results_table.rowCount()
                            self.results_table.insertRow(row)
                            self.results_table.setItem(
                                row, 0, QTableWidgetItem(os.path.basename(pdf_path))
                            )

                            status_item = QTableWidgetItem("Failed")
                            status_item.setData(Qt.UserRole, "failed")
                            self.results_table.setItem(row, 1, status_item)

                            # Validation Status - Set to "Pending" initially
                            validation_status_item = QTableWidgetItem("Pending")
                            validation_status_item.setData(Qt.UserRole, "pending")
                            validation_status_item.setForeground(QColor("#F59E0B"))  # Amber color for pending
                            self.results_table.setItem(row, 2, validation_status_item)

                            # Update failed counter
                            failed_count += 1
                            self.failed_count.setText(str(failed_count))

                            self.results_table.setItem(
                                row, 3, QTableWidgetItem(str(actual_page_count))
                            )  # Still show actual page count
                            self.results_table.setItem(row, 4, QTableWidgetItem("0"))
                            self.results_table.setItem(row, 5, QTableWidgetItem("0"))
                            self.results_table.setItem(row, 6, QTableWidgetItem("0"))

                    except Exception as e:
                        handle_exception(
                            func_name="process_files",
                            exception=e,
                            context={"pdf_path": pdf_path, "file_index": i}
                        )

                        # Try to get page count even if processing failed
                        try:
                            with fitz.open(pdf_path) as pdf:
                                actual_page_count = len(pdf)
                        except:
                            actual_page_count = 0

                        # Add error to results table
                        row = self.results_table.rowCount()
                        self.results_table.insertRow(row)
                        self.results_table.setItem(
                            row, 0, QTableWidgetItem(os.path.basename(pdf_path))
                        )

                        status_item = QTableWidgetItem(f"Error: {str(e)}")
                        status_item.setData(Qt.UserRole, "failed")
                        self.results_table.setItem(row, 1, status_item)

                        # Validation Status - Set to "Pending" initially
                        validation_status_item = QTableWidgetItem("Pending")
                        validation_status_item.setData(Qt.UserRole, "pending")
                        validation_status_item.setForeground(QColor("#F59E0B"))  # Amber color for pending
                        self.results_table.setItem(row, 2, validation_status_item)

                        # Update processed and failed counters
                        processed_count += 1
                        self.processed_count.setText(str(processed_count))
                        failed_count += 1
                        self.failed_count.setText(str(failed_count))

                        self.results_table.setItem(
                            row, 3, QTableWidgetItem(str(actual_page_count))
                        )
                        self.results_table.setItem(row, 4, QTableWidgetItem("0"))
                        self.results_table.setItem(row, 5, QTableWidgetItem("0"))
                        self.results_table.setItem(row, 6, QTableWidgetItem("0"))

                    # Update progress
                    self.progress_bar.setValue(batch_start + index + 1)
                    QApplication.processEvents()  # Keep UI responsive

            # Final update of processing time
            self.update_processing_time(is_final=True)

            # Hide stop button when done
            self.stop_button.setVisible(False)

            # Stop the processing timer
            self.processing_time_timer.stop()

            # Update license usage with the number of successfully processed files
            # Only count successful extractions towards the license limit
            self.update_license_usage(success_count)

            # Then update the color formatting for status items
            for row in range(self.results_table.rowCount()):
                status_item = self.results_table.item(row, 1)
                if status_item:
                    status_type = status_item.data(Qt.UserRole)
                    if status_type == "success":
                        status_item.setForeground(QColor(self.theme['secondary']))
                        status_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
                    elif status_type == "partial":
                        status_item.setForeground(QColor(self.theme['warning']))
                        status_item.setFont(QFont("Segoe UI", 9, QFont.Bold))
                    elif status_type == "failed":
                        status_item.setForeground(QColor(self.theme['danger']))
                        status_item.setFont(QFont("Segoe UI", 9, QFont.Bold))

            # Then modify the processing completion message and status labels to provide clearer information
            # Update status
            total_files = len(self.pdf_files)
            if success_count == total_files:
                self.status_label.setText("Processing complete: All files processed successfully!")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['secondary'] + '20'};
                    color: {self.theme['secondary']};
                    font-weight: bold;
                """)
                QMessageBox.information(self, "Success", f"All {total_files} files have been processed successfully.\nTotal time: {self.processing_time_label.text().replace('Total Time: ', '')}")
            elif success_count > 0:
                self.status_label.setText(f"Processing complete: {success_count}/{total_files} files processed successfully")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['warning'] + '20'};
                    color: {self.theme['warning']};
                    font-weight: bold;
                """)
                QMessageBox.warning(self, "Partial Success", f"{success_count} out of {total_files} files processed successfully.\n{failed_count} files failed.\nTotal time: {self.processing_time_label.text().replace('Total Time: ', '')}")
            else:
                self.status_label.setText("Processing complete: All files failed")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['danger'] + '20'};
                    color: {self.theme['danger']};
                    font-weight: bold;
                """)
                QMessageBox.critical(self, "Processing Failed", f"All {total_files} files failed to process. Please check logs for details.\nTotal time: {self.processing_time_label.text().replace('Total Time: ', '')}")

        except Exception as e:
            # Final update of processing time
            self.update_processing_time(is_final=True)

            # Hide stop button when done
            self.stop_button.setVisible(False)

            # Stop the processing timer
            if hasattr(self, 'processing_time_timer'):
                self.processing_time_timer.stop()

            print(f"Error in process_files: {str(e)}")
            import traceback

            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}\nTotal time: {self.processing_time_label.text().replace('Total Time: ', '')}")
            self.status_label.setText("Error occurred during processing")

    def add_files(self):
        """Add PDF files to the list"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF Files", "", "PDF Files (*.pdf)"
        )

        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.file_list.addItem(os.path.basename(file))

    def clear_files(self):
        """Clear the file list"""
        self.pdf_files.clear()
        self.file_list.clear()
        self.results_table.setRowCount(0)
        self.processed_data.clear()

    def export_data(self, section):
        """Export processed data in JSON format"""
        if not self.processed_data:
            QMessageBox.warning(
                self, "Warning", "No processed data available to export"
            )
            return

        try:
            # Create export directory if it doesn't exist
            export_dir = os.path.abspath("exported_data")
            os.makedirs(export_dir, exist_ok=True)

            # Update status
            self.status_label.setText(f"Exporting {section} data...")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['primary'] + '20'};
                color: {self.theme['primary']};
                font-weight: bold;
            """)
            QApplication.processEvents()  # Ensure UI updates

            # Prepare data for export
            export_data = {}

            # Get template ID for invoice2data processing
            template_id = self.get_selected_template_id()
            template_data = None

            if template_id:
                # Fetch template data for invoice2data processing
                try:
                    conn = sqlite3.connect("invoice_templates.db")
                    cursor = conn.cursor()
                    cursor.execute(
                        """
                        SELECT id, name, description, template_type, regions, column_lines, config, creation_date,
                               page_count, page_regions, page_column_lines, page_configs, json_template, extraction_method
                        FROM templates WHERE id = ?
                    """,
                        (template_id,),
                    )
                    template = cursor.fetchone()

                    if template:
                        template_data = {
                            "id": template[0],
                            "name": template[1],
                            "description": template[2],
                            "template_type": template[3],
                            "regions": json.loads(template[4]),
                            "column_lines": json.loads(template[5]),
                            "config": json.loads(template[6]),
                            "creation_date": template[7],
                            "page_count": template[8] if template[8] else 1,
                        }

                        # Load multi-page data if available
                        if template[9]:  # page_regions
                            template_data["page_regions"] = json.loads(template[9])

                        if template[10]:  # page_column_lines
                            template_data["page_column_lines"] = json.loads(template[10])

                        if template[11]:  # page_configs
                            template_data["page_configs"] = json.loads(template[11])

                        # Load JSON template for invoice2data if available
                        if template[12]:  # json_template
                            template_data["json_template"] = json.loads(template[12])

                        # Load extraction method if available
                        if len(template) > 13 and template[13]:  # extraction_method
                            template_data["extraction_method"] = template[13]
                        else:
                            template_data["extraction_method"] = "pypdf_table_extraction"  # default

                    conn.close()
                except Exception as e:
                    print(f"Error fetching template data: {str(e)}")
                    template_data = None

            for pdf_path, data in self.processed_data.items():
                pdf_filename = os.path.basename(pdf_path)
                template_type = data.get("template_type", "single")
                pdf_page_count = data.get("pdf_page_count", 1)

                # Create an entry for this PDF file
                file_data = {
                    "metadata": {
                        "filename": pdf_filename,
                        "page_count": pdf_page_count,
                        "template_type": template_type,
                        "export_date": datetime.now().isoformat(),
                        "template_name": self.template_combo.currentText(),
                    }
                }

                # Process with invoice2data if template data is available
                if template_data and "json_template" in template_data and template_data["json_template"]:
                    print(f"\nProcessing {pdf_filename} with invoice2data...")

                    # Use the current file_data as extracted data for invoice2data
                    invoice2data_result = self.process_with_invoice2data(pdf_path, template_data, file_data)

                    if invoice2data_result:
                        # Add invoice2data results to metadata
                        file_data["invoice2data"] = invoice2data_result
                        print(f"Added invoice2data results to export data")

                # Process section data based on the template type
                print(f"\nExporting {section} data for {pdf_filename}")

                # Check if section exists in data
                if section not in data:
                    print(f"  No {section} data found for this file")
                    file_data[section] = []
                    continue

                section_data = data[section]

                # Handle None or empty case
                if section_data is None:
                    print(f"  {section} data is None")
                    file_data[section] = []
                    continue

                # Handle case where data is a list of dataframes (multiple tables)
                if isinstance(section_data, list):
                    print(f"  Processing list of {len(section_data)} table(s)")
                    # Create a combined dictionary with table indexes
                    tables_dict = {}
                    valid_tables = 0

                    for i, df in enumerate(section_data):
                        try:
                            if df is None:
                                print(f"  Table {i} is None, skipping")
                                continue

                            # Convert string to DataFrame if needed
                            if isinstance(df, str):
                                print(f"  Table {i} is a string, converting to DataFrame")
                                df = pd.DataFrame([{"text": df}])

                            if df.empty:
                                print(f"  Table {i} is empty, skipping")
                                continue

                            valid_tables += 1
                            # Check if dataframe has page information
                            if "pdf_page" in df.columns:
                                print(f"  Table {i} has page information, grouping by page")
                                # Group by page
                                page_data = {}
                                for page_num, page_df in df.groupby("pdf_page"):
                                    page_num_int = int(page_num)
                                    page_df = page_df.drop(columns=["pdf_page"])
                                    page_data[f"page_{page_num_int}"] = page_df.to_dict(orient="records")
                                    print(f"    Page {page_num_int}: {len(page_df)} rows")
                                tables_dict[f"table_{i}"] = page_data
                            else:
                                # Single page data
                                print(f"  Table {i}: {len(df)} rows (no page info)")
                                tables_dict[f"table_{i}"] = df.to_dict(orient="records")
                        except Exception as e:
                            print(f"  Error processing table {i}: {str(e)}")
                            import traceback
                            traceback.print_exc()

                    print(f"  Processed {valid_tables} valid tables")
                    file_data[section] = tables_dict

                else:
                    # Regular case - single dataframe or string
                    try:
                        if isinstance(section_data, str):
                            print(f"  {section} data is a string, converting to DataFrame")
                            section_data = pd.DataFrame([{"text": section_data}])

                        if not hasattr(section_data, 'empty'):
                            print(f"  {section} data is not a DataFrame, converting")
                            # Try to convert to DataFrame if possible
                            try:
                                section_data = pd.DataFrame(section_data)
                            except:
                                print(f"  Cannot convert {section} data to DataFrame")
                                file_data[section] = [{"error": "Data format error"}]
                                continue

                        if section_data.empty:
                            print(f"  {section} DataFrame is empty")
                            file_data[section] = []
                            continue

                        rows = len(section_data)
                        cols = len(section_data.columns)
                        print(f"  {section} DataFrame has {rows} rows and {cols} columns")

                        # Check if multi-page processing is needed
                        if "pdf_page" in section_data.columns and template_type == "multi":
                            print(f"  Multi-page processing for {section}")
                            # Group by page
                            page_data = {}
                            for page_num, page_df in section_data.groupby("pdf_page"):
                                page_num_int = int(page_num)
                                page_df = page_df.drop(columns=["pdf_page"])
                                page_data[f"page_{page_num_int}"] = page_df.to_dict(orient="records")
                                print(f"    Page {page_num_int}: {len(page_df)} rows")
                            file_data[section] = page_data
                        else:
                            # Single page data
                            if "pdf_page" in section_data.columns:
                                print(f"  Removing pdf_page column")
                                section_data = section_data.drop(columns=["pdf_page"])
                            print(f"  Exporting as single-page data: {len(section_data)} rows")
                            file_data[section] = section_data.to_dict(orient="records")
                    except Exception as e:
                        print(f"  Error processing {section} data: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        file_data[section] = [{"error": str(e)}]

                # Add the file data to the export
                export_data[pdf_filename] = file_data

            # Save to JSON file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(export_dir, f"{section}_data_{timestamp}.json")

            print(f"Attempting to save file to: {filename}")

            with open(filename, "w", encoding="utf-8") as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False, default=str)

            # Verify the file was created
            if not os.path.exists(filename):
                print(f"Warning: Failed to create file {filename}")

            # Reset status label to normal
            self.status_label.setText(f"Exported {section} data successfully")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['secondary'] + '20'};
                color: {self.theme['secondary']};
                font-weight: bold;
            """)

            # Create a custom success message box
            success_box = QMessageBox(self)
            success_box.setWindowTitle("Export Successful")
            success_box.setIcon(QMessageBox.Information)

            # Calculate total rows exported
            total_exported_rows = 0
            total_exported_files = len(export_data)

            for file_data in export_data.values():
                section_content = file_data.get(section, {})
                if isinstance(section_content, list):
                    total_exported_rows += len(section_content)
                elif isinstance(section_content, dict):
                    for table_data in section_content.values():
                        if isinstance(table_data, list):
                            total_exported_rows += len(table_data)
                        elif isinstance(table_data, dict):
                            for page_data in table_data.values():
                                if isinstance(page_data, list):
                                    total_exported_rows += len(page_data)

            success_box.setText(f"Data exported successfully to")
            success_box.setInformativeText(
                f"<b>File:</b> {filename}<br><br>"
                f"<b>Export details:</b><br>"
                f"• Section: <b>{section.title()}</b><br>"
                f"• Files: <b>{total_exported_files}</b><br>"
                f"• Rows: <b>{total_exported_rows}</b><br>"
            )

            # Verify the export directory exists
            os.makedirs(export_dir, exist_ok=True)

            # Verify the file exists
            if os.path.exists(filename):
                print(f"File exists and will be opened: {filename}")

                # Open folder button
                open_folder_btn = success_box.addButton("Open Folder", QMessageBox.ActionRole)
                # Use a direct function instead of a lambda to avoid potential issues
                def open_folder():
                    try:
                        folder_path = os.path.dirname(filename)
                        print(f"Opening folder: {folder_path}")
                        os.startfile(folder_path)
                    except Exception as e:
                        print(f"Error opening folder: {str(e)}")
                        QMessageBox.warning(self, "Error", f"Could not open folder: {str(e)}")

                open_folder_btn.clicked.connect(open_folder)

                # Open file button
                open_file_btn = success_box.addButton("Open File", QMessageBox.ActionRole)
                # Use a direct function instead of a lambda to avoid potential issues
                def open_file():
                    try:
                        print(f"Opening file: {filename}")
                        os.startfile(filename)
                    except Exception as e:
                        print(f"Error opening file: {str(e)}")
                        QMessageBox.warning(self, "Error", f"Could not open file: {str(e)}")

                open_file_btn.clicked.connect(open_file)
            else:
                print(f"Warning: Export file {filename} does not exist")

            # OK button
            ok_btn = success_box.addButton(QMessageBox.Ok)
            ok_btn.setDefault(True)

            success_box.exec()

        except Exception as e:
            self.status_label.setText("Export error")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['danger'] + '20'};
                color: {self.theme['danger']};
                font-weight: bold;
            """)

            QMessageBox.critical(self, "Error", f"Failed to export data: {str(e)}")
            import traceback

            traceback.print_exc()

    def navigate_back(self):
        """Return to the main screen"""
        self.go_back.emit()  # Emit the signal for parent to handle
        print("Emitted go_back signal")

    # Admin dialog removed for security

    def reset_screen(self):
        """Reset the screen to its initial state, clear cache, and clean temp directory"""
        # Clear all data
        self.pdf_files.clear()
        self.file_list.clear()
        self.results_table.setRowCount(0)
        self.processed_data.clear()

        # Reset progress bar
        self.progress_bar.setValue(0)

        # Reset status label
        self.status_label.setText("Ready")

        # Reset processing time label
        self.processing_time_label.setText("")
        self.start_time = None

        # Stop any running timer
        if hasattr(self, 'processing_time_timer') and self.processing_time_timer.isActive():
            self.processing_time_timer.stop()

        # Reset extraction statistics summary values
        self.processed_count.setText("0")
        self.success_count.setText("0")
        self.failed_count.setText("0")
        self.total_rows_count.setText("0")

        # Reset template selection to the placeholder item
        if self.template_combo.count() > 0:
            self.template_combo.setCurrentIndex(0)  # Select the "-- Select a template --" item

        # Get cache stats before clearing
        cache_stats = get_extraction_cache_stats()
        cache_size_before = cache_stats.get('extraction_cache_size', 0) + cache_stats.get('multipage_cache_size', 0)

        # Clear memory cache
        clear_extraction_cache()

        # Update cache size display
        self.cache_size_label.setText("0 items")

        # Force aggressive garbage collection - run multiple collection cycles
        for _ in range(3):
            gc.collect()

        # Release as much memory as possible
        import ctypes
        if hasattr(ctypes, 'windll'):  # Windows only
            try:
                ctypes.windll.kernel32.SetProcessWorkingSetSize(-1, -1)
                print("Released process working set memory")
            except Exception as e:
                print(f"Error releasing process memory: {str(e)}")

        # Small delay to allow memory to be properly released
        QTimer.singleShot(100, self.update_memory_usage)

        # Update memory usage immediately as well
        self.update_memory_usage()

        # Log the cache cleanup
        print(f"Cleared extraction cache ({cache_size_before} items)")

        # Clean up temporary directory
        temp_files_removed = 0
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                # Count files before cleanup
                for root, dirs, files in os.walk(self.temp_dir):
                    temp_files_removed += len(files)
                    temp_files_removed += len(dirs)

                # Clean up the temp directory
                self.cleanup_temp_directory()

                # Log the temp directory cleanup
                print(f"Cleaned temporary directory: {self.temp_dir} ({temp_files_removed} files/folders removed)")
            except Exception as e:
                print(f"Error cleaning temporary directory: {str(e)}")
                import traceback
                traceback.print_exc()

        # Update status message
        if temp_files_removed > 0:
            self.status_label.setText(f"Cleared cache and temp files ({cache_size_before} items, {temp_files_removed} temp files)")
        else:
            self.status_label.setText(f"Cleared cache ({cache_size_before} items)")

        # Reset status after a short delay
        QTimer.singleShot(3000, lambda: self.status_label.setText("Ready"))

        # Reload license information
        self.load_license_info()

        # Show confirmation message
        message = f"The screen has been reset to its initial state.\n\n"
        message += f"• Extraction cache cleared: {cache_size_before} items\n"
        if temp_files_removed > 0:
            message += f"• Temporary files removed: {temp_files_removed} files/folders"

        QMessageBox.information(
            self, "Screen Reset", message
        )

    def start_memory_monitoring(self):
        """Start monitoring memory usage"""
        # Initialize memory monitoring timer
        self.memory_monitor_timer = QTimer(self)
        self.memory_monitor_timer.timeout.connect(self.update_memory_usage)
        self.memory_monitor_timer.start(5000)  # Update every 5 seconds

        # Initial update
        self.update_memory_usage()

    def stop_memory_monitoring(self):
        """Stop monitoring memory usage"""
        if self.memory_monitor_timer and self.memory_monitor_timer.isActive():
            self.memory_monitor_timer.stop()

    def update_memory_usage(self):
        """Update memory usage display"""
        try:
            # Get current process
            process = psutil.Process(os.getpid())

            # Get memory info
            memory_info = process.memory_info()
            memory_usage_bytes = memory_info.rss  # Resident Set Size in bytes

            # Convert to MB for display
            memory_usage_mb = memory_usage_bytes / (1024 * 1024)

            # Get system memory info
            system_memory = psutil.virtual_memory()
            total_memory_mb = system_memory.total / (1024 * 1024)

            # Calculate percentage
            memory_percentage = (memory_usage_bytes / system_memory.total) * 100
            self.current_memory_usage = memory_percentage

            # Update display
            self.memory_usage_label.setText(f"{memory_percentage:.1f}% ({memory_usage_mb:.1f} MB)")

            # Set color based on usage
            if memory_percentage > 80:
                self.memory_usage_label.setStyleSheet(f"color: {self.theme['danger']}; font-weight: bold;")
            elif memory_percentage > 60:
                self.memory_usage_label.setStyleSheet(f"color: {self.theme['warning']}; font-weight: bold;")
            else:
                self.memory_usage_label.setStyleSheet(f"color: {self.theme['tertiary']}; font-weight: bold;")

            # Update cache size display
            cache_stats = get_extraction_cache_stats()
            total_cache_size = cache_stats.get('extraction_cache_size', 0) + cache_stats.get('multipage_cache_size', 0)
            self.cache_size_label.setText(f"{total_cache_size} items")

        except Exception as e:
            handle_exception(
                func_name="update_memory_usage",
                exception=e,
                context={"memory_monitoring": True}
            )
            self.memory_usage_label.setText("Error")
            self.cache_size_label.setText("Error")

    def check_memory_usage(self):
        """Check memory usage and clean up if necessary"""
        self.update_memory_usage()

        # If memory usage exceeds threshold, clean up
        if self.current_memory_usage > self.memory_threshold:
            print(f"Memory usage ({self.current_memory_usage:.1f}%) exceeds threshold ({self.memory_threshold}%). Cleaning up...")

            try:
                # Get cache stats before clearing
                cache_stats = get_extraction_cache_stats()
                cache_size_before = cache_stats.get('extraction_cache_size', 0) + cache_stats.get('multipage_cache_size', 0)

                # Clear the cache
                clear_extraction_cache()

                # Update cache size display
                self.cache_size_label.setText("0 items")

                # Log the cleanup
                print(f"Cleared extraction cache ({cache_size_before} items)")

                # Show a brief status message
                self.status_label.setText(f"Cleared cache ({cache_size_before} items)")

                # Reset status after a short delay
                QTimer.singleShot(3000, lambda: self.status_label.setText("Ready"))

            except Exception as e:
                print(f"Error clearing memory cache: {str(e)}")
                import traceback
                traceback.print_exc()

            # Force aggressive garbage collection - run multiple collection cycles
            for _ in range(3):
                gc.collect()

            # Release as much memory as possible
            import ctypes
            if hasattr(ctypes, 'windll'):  # Windows only
                try:
                    ctypes.windll.kernel32.SetProcessWorkingSetSize(-1, -1)
                    print("Released process working set memory")
                except Exception as e:
                    print(f"Error releasing process memory: {str(e)}")

            # Small delay to allow memory to be properly released
            QTimer.singleShot(100, self.update_memory_usage)

            # Update memory usage immediately as well
            self.update_memory_usage()



    def setup_temp_directory(self):
        """Create a temporary directory for invoice files and templates"""
        try:
            # Clean up any existing temp directory
            self.cleanup_temp_directory()

            # Create a new temporary directory
            self.temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
            os.makedirs(self.temp_dir, exist_ok=True)

            # Register with cache manager for cleanup
            if CACHE_MANAGER_AVAILABLE:
                try:
                    register_temp_directory(self.temp_dir)
                    print(f"[DEBUG] Registered temp directory with cache manager: {self.temp_dir}")
                except Exception as e:
                    print(f"[WARNING] Failed to register temp directory with cache manager: {e}")

            print(f"Created temporary directory: {self.temp_dir}")
            return self.temp_dir
        except Exception as e:
            print(f"Error creating temporary directory: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def cleanup_temp_directory(self):
        """Clean up the temporary directory"""
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                # Clear the invoice txt files mapping
                self.invoice_txt_files = {}
                self.template_file_path = None

                # Remove all files in the temp directory
                for filename in os.listdir(self.temp_dir):
                    file_path = os.path.join(self.temp_dir, filename)
                    try:
                        if os.path.isfile(file_path):
                            os.unlink(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print(f"Error removing {file_path}: {str(e)}")

                print(f"Cleaned temporary directory: {self.temp_dir}")
        except Exception as e:
            print(f"Error cleaning temporary directory: {str(e)}")
            import traceback
            traceback.print_exc()

    def cleanup_bulk_processor_caches(self):
        """Clean up all caches and temporary data for the bulk processor"""
        try:
            print("[DEBUG] Starting bulk_processor cache cleanup...")

            # Clear extraction caches
            clear_extraction_cache()
            print("[DEBUG] Cleared all extraction caches")

            # Clear processed data
            if hasattr(self, 'processed_data'):
                self.processed_data.clear()
                print("[DEBUG] Cleared processed data")

            # Clear PDF files list
            if hasattr(self, 'pdf_files'):
                self.pdf_files.clear()
                print("[DEBUG] Cleared PDF files list")

            # Clean up temporary directory
            self.cleanup_temp_directory()

            # Clear invoice txt files mapping
            if hasattr(self, 'invoice_txt_files'):
                self.invoice_txt_files.clear()
                print("[DEBUG] Cleared invoice txt files mapping")

            # Clear template file path
            if hasattr(self, 'template_file_path'):
                self.template_file_path = None
                print("[DEBUG] Cleared template file path")

            # Force garbage collection
            import gc
            collected = gc.collect()
            print(f"[DEBUG] Garbage collection: {collected} objects collected")

            print("[DEBUG] Bulk processor cache cleanup completed")

        except Exception as e:
            print(f"[ERROR] Error in cleanup_bulk_processor_caches: {str(e)}")
            import traceback
            traceback.print_exc()

    def convert_extraction_to_text(self, pdf_path, extracted_data, for_display=False):
        """Convert the extraction data to a text format that can be used by invoice2data

        Args:
            pdf_path (str): Path to the PDF file
            extracted_data (dict): The extracted data from pypdf_table_extraction
            for_display (bool): If True, returns clean format without pipe separators

        Returns:
            str: Text representation of the extracted data with pipe-separated values
        """
        print(f"[DEBUG] Converting extraction to text for {pdf_path} (for_display={for_display})")

        if not extracted_data:
            print(f"[DEBUG] No extracted data for {pdf_path}")
            return ""

        # Use the shared utility function from invoice2data_utils
        return invoice2data_utils.convert_extraction_to_text(extracted_data, pdf_path=pdf_path, for_display=for_display)

    def clean_dataframe(self, df, section, config):
        """Basic DataFrame cleaning without regex patterns"""
        if df is None or df.empty:
            return df

        print(f"Cleaning {section} DataFrame with basic cleaning")

        # Basic cleaning - replace empty strings with NaN but don't remove empty rows/columns
        df = df.replace(r"^\s*$", pd.NA, regex=True)

        # No longer dropping empty rows and columns
        # df = df.dropna(how="all")
        # df = df.dropna(axis=1, how="all")

        # Clean string values
        for col in df.columns:
            if df[col].dtype == object:  # Only clean string columns
                df[col] = df[col].apply(
                    lambda x: x.strip() if isinstance(x, str) else x
                )

        print(f"  Final DataFrame size: {len(df)} rows, {len(df.columns)} columns")
        return df

    def extract_invoice_tables(self, pdf_path, template_id):
        try:
            print("\n" + "=" * 80)
            print(f"STEP 1: DATABASE CONNECTION AND TEMPLATE RETRIEVAL")
            print("=" * 80)

            # Connect to database
            print("Connecting to database: 'invoice_templates.db'")
            conn = sqlite3.connect("invoice_templates.db")
            cursor = conn.cursor()

            # Fetch template data including dual coordinate columns and extraction method
            cursor.execute(
                """
                SELECT id, name, description, template_type, regions, column_lines, config, creation_date,
                       page_count, page_regions, page_column_lines, page_configs, json_template,
                       drawing_regions, drawing_column_lines, extraction_regions, extraction_column_lines,
                       drawing_page_regions, drawing_page_column_lines, extraction_page_regions, extraction_page_column_lines,
                       extraction_method
                FROM templates WHERE id = ?
            """,
                (template_id,),
            )
            template = cursor.fetchone()

            if not template:
                raise Exception(f"Template with ID {template_id} not found")

            # Extract template data with dual coordinate support
            from dual_coordinate_storage import DualCoordinateStorage
            storage = DualCoordinateStorage()

            template_data = {
                "id": template[0],
                "name": template[1],
                "description": template[2],
                "template_type": template[3],
                "regions": json.loads(template[4]) if template[4] else {'header': [], 'items': [], 'summary': []},  # Legacy
                "column_lines": json.loads(template[5]) if template[5] else {'header': [], 'items': [], 'summary': []},  # Legacy
                "config": json.loads(template[6]),
                "creation_date": template[7],
                "page_count": template[8] if template[8] else 1,
            }

            # Load multi-page data if available
            if template[9]:  # page_regions
                template_data["page_regions"] = json.loads(template[9])

            if template[10]:  # page_column_lines
                template_data["page_column_lines"] = json.loads(template[10])

            if template[11]:  # page_configs
                template_data["page_configs"] = json.loads(template[11])

            # Load dual coordinate data if available
            if len(template) > 13:  # Check if dual coordinate columns exist
                if template[13]:  # drawing_regions
                    template_data["drawing_regions"] = storage.deserialize_regions(template[13])
                if template[14]:  # drawing_column_lines
                    template_data["drawing_column_lines"] = storage.deserialize_column_lines(template[14])
                if template[15]:  # extraction_regions
                    template_data["extraction_regions"] = storage.deserialize_regions(template[15])
                if template[16]:  # extraction_column_lines
                    template_data["extraction_column_lines"] = storage.deserialize_column_lines(template[16])
                if template[17]:  # drawing_page_regions
                    template_data["drawing_page_regions"] = json.loads(template[17]) if template[17] else []
                if template[18]:  # drawing_page_column_lines
                    template_data["drawing_page_column_lines"] = json.loads(template[18]) if template[18] else []
                if template[19]:  # extraction_page_regions
                    template_data["extraction_page_regions"] = json.loads(template[19]) if template[19] else []
                if template[20]:  # extraction_page_column_lines
                    template_data["extraction_page_column_lines"] = json.loads(template[20]) if template[20] else []

            # Load JSON template for invoice2data if available
            if template[12]:  # json_template
                template_data["json_template"] = json.loads(template[12])

            # Load extraction method if available (this is from the updated query above)
            extraction_method = "pypdf_table_extraction"  # default
            if len(template) > 21:  # Check if extraction_method column exists (index 21 in the updated query)
                # The extraction_method should be at index 21 in the updated query
                if template[21]:
                    extraction_method = template[21]
            template_data["extraction_method"] = extraction_method
            print(f"Using extraction method: {extraction_method}")

            # Close database connection
            conn.close()

            # Check if dual coordinate data is available and use it for extraction
            has_dual_coords = 'extraction_regions' in template_data or 'extraction_page_regions' in template_data
            print(f"Dual coordinate data available: {has_dual_coords}")

            if has_dual_coords:
                print("Using dual coordinate system for extraction - no coordinate conversion needed")
                return self._extract_with_dual_coordinates(pdf_path, template_data)
            else:
                print("Using legacy coordinate system - coordinate conversion required")

            print("\n" + "=" * 80)
            print(f"STEP 2: PDF DOCUMENT LOADING")
            print("=" * 80)

            # Load the PDF document
            print(f"Loading PDF document: {pdf_path}")
            pdf_document = fitz.open(pdf_path)
            pdf_page_count = len(pdf_document)
            print(f"✓ PDF document loaded successfully with {pdf_page_count} pages")

            print("\n" + "=" * 80)
            print(f"STEP 3: TABLE EXTRACTION")
            print("=" * 80)

            # Initialize results dictionary with extraction statuses
            results = {
                "header_tables": [],
                "items_tables": [],
                "summary_tables": [],
                "extraction_status": {
                    "header": "not_processed",
                    "items": "not_processed",
                    "summary": "not_processed",
                    "overall": "not_processed"
                }
            }

            # Get config parameters
            config = template_data.get("config", {})

            # Create a template manager instance to handle middle page logic
            from template_manager import TemplateManager
            template_mgr = TemplateManager()

            # Multi-page options removed - using simplified page-wise approach
            # Page mapping features will be implemented later as a separate enhancement

            if template_data["template_type"] == "multi":
                print(f"Multi-page template detected - using simplified page-wise approach")

            # Determine which pages to process
            if template_data["template_type"] == "single":
                # For single-page templates, only process the first page
                pages_to_process = [0]  # First page
            else:
                # For multi-page templates, process all pages
                pages_to_process = list(range(pdf_page_count))

            print(f"Pages to process: {[p+1 for p in pages_to_process]}")

            # Use the unified invoice processing utilities to get the page mapping
            page_mapping = invoice_processing_utils.apply_template_with_middle_page_logic(
                template_data=template_data,
                pdf_path=pdf_path,
                pdf_total_pages=pdf_page_count
            )

            # Process each selected page
            for page_index in pages_to_process:
                print(f"\nProcessing page {page_index + 1}/{pdf_page_count}")

                try:
                    # Get the current page
                    page = pdf_document[page_index]

                    # Get regions and column lines for current page using the template manager mapping
                    if page_index in page_mapping:
                        page_data = page_mapping[page_index]
                        current_regions = page_data['regions']
                        current_column_lines = page_data['column_lines']
                        template_page_idx = page_data['template_page_idx']

                        print(f"Using template page {template_page_idx + 1} for PDF page {page_index + 1}")
                    else:
                        print(f"Warning: No template mapping for page {page_index + 1}")
                        continue

                    # Debug column lines for templates
                    print(f"\nTemplate column lines:")
                    for section, lines in current_column_lines.items():
                        print(f"  {section}: {len(lines)} column lines")
                        if lines and len(lines) > 0:
                            print(f"    First column line format: {type(lines[0])}")
                            print(f"    Sample: {lines[0]}")

                    # Verify column_lines structure - it should be a dict with sections as keys
                    if not isinstance(current_column_lines, dict):
                        print(f"  WARNING: column_lines is not a dict: {type(current_column_lines)}")
                        # Try to fix it - common issue is an array with a single entry
                        if isinstance(current_column_lines, list) and len(current_column_lines) > 0:
                            print(f"  Attempting to fix column_lines format (found list with {len(current_column_lines)} entries)")
                            # Use the first item if it's a dict
                            if isinstance(current_column_lines[0], dict):
                                current_column_lines = current_column_lines[0]
                                print(f"  Fixed column_lines to use first entry: {type(current_column_lines)}")

                    # Debug the regions we're using
                    print(f"Using regions format: {type(current_regions)}")
                    if current_regions:
                        for section, regions in current_regions.items():
                            print(f"  {section}: {len(regions)} region(s)")
                            if regions:
                                print(f"    First region type: {type(regions[0])}")

                    # Process each section (header, items, summary)
                    for section in ["header", "items", "summary"]:
                        if section in current_regions and current_regions[section]:
                            section_regions = current_regions[section]
                            section_column_lines = current_column_lines.get(section, [])

                            print(
                                f"\nExtracting {section} section from page {page_index + 1}"
                            )
                            print(f"  Found {len(section_regions)} region(s)")
                            print(f"  Found {len(section_column_lines)} column line(s)")

                            # Ensure section_column_lines is a list
                            if not isinstance(section_column_lines, list):
                                print(f"  WARNING: section_column_lines is not a list: {type(section_column_lines)}")
                                if isinstance(section_column_lines, dict):
                                    # Try to convert dict to list
                                    section_column_lines = [section_column_lines]
                                    print(f"  Converted dict to list with 1 item")
                                else:
                                    # Initialize as empty list as fallback
                                    section_column_lines = []
                                    print(f"  Reset to empty list as fallback")

                            # Get table extraction parameters
                            table_areas = []
                            columns_list = []

                            for region_idx, region in enumerate(section_regions):
                                # Handle different region formats
                                if isinstance(region, dict):
                                    # Format from single page viewer: {x1, y1, x2, y2}
                                    x1 = region.get("x1", 0)
                                    y1 = region.get("y1", 0)
                                    x2 = region.get("x2", 0)
                                    y2 = region.get("y2", 0)
                                elif isinstance(region, list) and len(region) >= 2:
                                    # Format from multi-page viewer: [{x,y}, {x,y}]
                                    x1 = region[0].get("x", 0)
                                    y1 = region[0].get("y", 0)
                                    x2 = region[1].get("x", 0)
                                    y2 = region[1].get("y", 0)
                                else:
                                    print(f"  Warning: Unrecognized region format: {region}")
                                    continue

                                # Create table area string
                                table_area = f"{x1},{y1},{x2},{y2}"
                                table_areas.append(table_area)

                                # Process column lines for this region
                                region_columns = []

                                # Debug all column lines for this section
                                print(f"  Processing column lines for region {region_idx} in {section} section")
                                print(f"  Column lines count: {len(section_column_lines)}")
                                print(f"  Column lines type: {type(section_column_lines)}")

                                # Fix for single page templates: check structure of column lines
                                if len(section_column_lines) == 0 and template_data["template_type"] == "single":
                                    print(f"  WARNING: No column lines found for {section} section in single page template")
                                    print(f"  Template data column_lines structure: {type(template_data.get('column_lines', {}))}")

                                    # Try to directly access the column lines from the template data
                                    all_column_lines = template_data.get("column_lines", {})
                                    if isinstance(all_column_lines, dict) and section in all_column_lines:
                                        direct_section_column_lines = all_column_lines.get(section, [])
                                        if direct_section_column_lines:
                                            print(f"  Found {len(direct_section_column_lines)} column lines directly in template data")
                                            section_column_lines = direct_section_column_lines
                                            print(f"  First entry type: {type(direct_section_column_lines[0])}")

                                    # Also check if column_lines might be an array itself (format inconsistency)
                                    if isinstance(all_column_lines, list) and len(all_column_lines) > 0:
                                        print(f"  Column lines is a list with {len(all_column_lines)} entries")
                                        # Use the first entry for single page templates
                                        if isinstance(all_column_lines[0], dict) and section in all_column_lines[0]:
                                            first_page_column_lines = all_column_lines[0].get(section, [])
                                            if first_page_column_lines:
                                                print(f"  Found {len(first_page_column_lines)} column lines in first page entry")
                                                section_column_lines = first_page_column_lines

                                for line in section_column_lines:
                                    # Handle different column line formats
                                    if isinstance(line, list):
                                        if len(line) >= 3 and line[2] == region_idx:
                                            x_val = line[0].get("x", 0)
                                            region_columns.append(x_val)
                                            print(f"    Using column at x={x_val} (matched region_idx={region_idx})")
                                        elif len(line) == 2:
                                            # Format without region index: [{x,y}, {x,y}]
                                            x_val = line[0].get("x", 0)
                                            region_columns.append(x_val)
                                            print(f"    Using column at x={x_val} (list format)")
                                    elif isinstance(line, dict):
                                        # Try different known formats
                                        if "x" in line:
                                            # Direct format: {x, y}
                                            x_val = line.get("x", 0)
                                            region_columns.append(x_val)
                                            print(f"    Using column at x={x_val} (dict format with x key)")
                                        elif "x1" in line:
                                            # Rectangle format: {x1, y1, x2, y2}
                                            x_val = line.get("x1", 0)
                                            region_columns.append(x_val)
                                            print(f"    Using column at x={x_val} (dict format with x1 key)")
                                        elif "value" in line and isinstance(line["value"], (int, float)):
                                            # Value format: {value: 123}
                                            x_val = line["value"]
                                            region_columns.append(x_val)
                                            print(f"    Using column at x={x_val} (dict format with value key)")
                                        elif "position" in line:
                                            # Position format: {position: 123}
                                            x_val = line["position"]
                                            region_columns.append(x_val)
                                            print(f"    Using column at x={x_val} (dict format with position key)")
                                        else:
                                            # Unknown dict format, try to extract any numeric value
                                            print(f"    Unknown dict format: {line}")
                                            for key, val in line.items():
                                                if isinstance(val, (int, float)):
                                                    print(f"    Using numeric value {val} from key '{key}'")
                                                    region_columns.append(val)
                                                    break
                                    elif isinstance(line, (int, float)):
                                        # Direct numeric value
                                        region_columns.append(line)
                                        print(f"    Using direct numeric value: x={line}")
                                    else:
                                        print(f"    Unsupported column line format: {type(line)}, value: {line}")
                                        # Try to extract a numeric value if it's a string
                                        if isinstance(line, str):
                                            try:
                                                numeric_val = float(line)
                                                region_columns.append(numeric_val)
                                                print(f"    Converted string to numeric value: x={numeric_val}")
                                            except ValueError:
                                                print(f"    Could not convert string to numeric value")

                                # Format column lines
                                col_str = (
                                    ",".join([str(x) for x in sorted(region_columns)])
                                    if region_columns
                                    else ""
                                )
                                columns_list.append(col_str)

                            # Handle special case for items section with multiple regions
                            if section == "items" and len(table_areas) > 1:
                                print(f"  Combining multiple item regions into one ({len(table_areas)} regions)")

                                # Parse all coordinates
                                area_coords = []
                                for area in table_areas:
                                    coords = [float(c) for c in area.split(",")]
                                    area_coords.append(coords)

                                # Find bounding box
                                x_coords = [c[0] for c in area_coords] + [
                                    c[2] for c in area_coords
                                ]
                                y_coords = [c[1] for c in area_coords] + [
                                    c[3] for c in area_coords
                                ]

                                x1 = min(x_coords)
                                y1 = min(y_coords)
                                x2 = max(x_coords)
                                y2 = max(y_coords)

                                # Replace with combined area
                                combined_area = f"{x1},{y1},{x2},{y2}"
                                table_areas = [combined_area]
                                print(f"  Combined area: {combined_area}")

                                # Combine column lines with deduplication
                                all_columns = set()  # Use set for deduplication
                                for col_str in columns_list:
                                    if col_str:
                                        for col in col_str.split(","):
                                            all_columns.add(float(col))

                                # Convert back to list and sort
                                all_columns_list = sorted(list(all_columns))

                                # Remove columns that are too close to each other (within 5 pixels)
                                if len(all_columns_list) > 1:
                                    deduplicated_columns = [all_columns_list[0]]
                                    for i in range(1, len(all_columns_list)):
                                        if all_columns_list[i] - deduplicated_columns[-1] >= 5:  # 5 pixel threshold
                                            deduplicated_columns.append(all_columns_list[i])
                                    all_columns_list = deduplicated_columns

                                # Format as string
                                col_str = ",".join([str(x) for x in all_columns_list]) if all_columns_list else ""
                                columns_list = [col_str]
                                print(f"  Combined columns: {col_str}")

                            # Set extraction parameters from config
                            extraction_params = {}
                            # Check if we have extraction_params in the config
                            if "extraction_params" in config:
                                extraction_params = config["extraction_params"]
                                print(f"  Found extraction_params in config")
                            else:
                                print(
                                    f"  WARNING: No extraction_params found in config, using direct config values"
                                )

                            # Get section-specific parameters
                            section_params = {}
                            if section in extraction_params:
                                section_params = extraction_params.get(section, {})
                                print(
                                    f"  Found section-specific parameters for {section}"
                                )

                            # Extract row_tol with proper fallbacks
                            row_tol = section_params.get("row_tol", None)
                            if row_tol is not None:
                                print(
                                    f"  Using row_tol={row_tol} from extraction_params.{section}"
                                )
                            else:
                                # If not in section_params, check direct section config
                                section_config = config.get(section, {})
                                row_tol = section_config.get("row_tol", None)
                                if row_tol is not None:
                                    print(
                                        f"  Using row_tol={row_tol} from config.{section}"
                                    )
                                else:
                                    # If not in direct section config, check global config
                                    row_tol = config.get("row_tol", None)
                                if row_tol is not None:
                                    print(f"  Using global row_tol={row_tol}")
                                else:
                                    # STRICT MODE: Raise error instead of using default
                                    error_msg = f"ERROR: No row_tol defined for {section} in database config"
                                    print(f"  ❌ {error_msg}")
                                    raise ValueError(error_msg)

                            # Get other extraction parameters with fallbacks
                            split_text = extraction_params.get(
                                "split_text", config.get("split_text", True)
                            )
                            strip_text = extraction_params.get(
                                "strip_text", config.get("strip_text", "\n")
                            )
                            flavor = extraction_params.get(
                                "flavor", config.get("flavor", "stream")
                            )
                            print(f"  Extraction parameters for {section}:")
                            print(f"    Table areas: {table_areas}")
                            print(f"    Columns: {columns_list}")
                            print(f"    Row tolerance: {row_tol} (from database)")

                            # Extract tables for this section using standardized approach
                            try:
                                # Prepare raw extraction parameters
                                raw_extraction_params = {
                                    section: {"row_tol": row_tol},
                                    "split_text": split_text,
                                    "strip_text": strip_text,
                                    "flavor": flavor,
                                    "parallel": True
                                }

                                # Normalize extraction parameters
                                normalized_params = normalize_extraction_params(raw_extraction_params)

                                # Convert table_areas and columns_list to proper format
                                processed_table_areas = []
                                processed_columns_list = []

                                for table_area, columns in zip(table_areas, columns_list):
                                    # Convert table_area from string to list of floats if needed
                                    if isinstance(table_area, str):
                                        table_area = [float(x) for x in table_area.split(',')]
                                    processed_table_areas.append(table_area)

                                    # Convert columns from string to list of floats if needed
                                    if columns and isinstance(columns, str):
                                        columns = [float(x) for x in columns.split(',') if x.strip()]
                                    processed_columns_list.append(columns)

                                # Check if we have valid table areas (not all zeros)
                                valid_areas = [area for area in processed_table_areas if area != [0.0, 0.0, 0.0, 0.0]]

                                if not valid_areas:
                                    print(f"  WARNING: No valid table areas found for {section} section (all coordinates are 0,0,0,0)")
                                    print(f"  This usually means the template regions are not properly configured")
                                    continue

                                # Use multi-method extraction logic for valid areas
                                extraction_method = template_data.get("extraction_method", "pypdf_table_extraction")

                                if len(valid_areas) > 1:
                                    # Multiple tables - use multi-method extraction
                                    table_dfs = extract_with_method(
                                        pdf_path=pdf_path,
                                        extraction_method=extraction_method,
                                        page_number=page_index + 1,
                                        table_areas=valid_areas,
                                        columns_list=processed_columns_list[:len(valid_areas)],
                                        section_type=section,
                                        extraction_params=normalized_params,
                                        use_cache=True
                                    )

                                    # Process each table
                                    for table_index, table_df in enumerate(table_dfs or []):
                                        if table_df is not None and not table_df.empty:
                                            self._process_extracted_table(
                                                table_df, section, results,
                                                region_index=table_index,
                                                page_number=page_index + 1
                                            )
                                else:
                                    # Single table - use multi-method extraction
                                    table_df = extract_with_method(
                                        pdf_path=pdf_path,
                                        extraction_method=extraction_method,
                                        page_number=page_index + 1,
                                        table_areas=valid_areas,
                                        columns_list=processed_columns_list if processed_columns_list else None,
                                        section_type=section,
                                        extraction_params=normalized_params,
                                        use_cache=True
                                    )

                                    if table_df is not None and not table_df.empty:
                                        self._process_extracted_table(
                                            table_df, section, results,
                                            region_index=0,  # Single table, so index 0
                                            page_number=page_index + 1
                                        )

                            except Exception as e:
                                handle_exception(
                                    func_name="extract_tables_for_section",
                                    exception=e,
                                    context={"section": section, "page": page_index + 1}
                                )
                        else:
                            log_info(f"No {section} regions defined for page {page_index + 1}")

                except Exception as e:
                    handle_exception(
                        func_name="process_page",
                        exception=e,
                        context={"page": page_index + 1, "pdf_path": pdf_path}
                    )

            # Close the PDF document
            pdf_document.close()

            # At the end of processing all pages, update the overall extraction status
            # Update the overall extraction status before returning results
            if results["extraction_status"]["items"] == "success":
                # If items were successfully extracted, that's most important
                results["extraction_status"]["overall"] = "success"
            elif results["extraction_status"]["items"] == "partial" or results["extraction_status"]["header"] == "success" or results["extraction_status"]["summary"] == "success":
                # Partial success if we at least got some data
                results["extraction_status"]["overall"] = "partial"
            else:
                # Failed if nothing was successfully extracted
                results["extraction_status"]["overall"] = "failed"

            print(f"\nExtraction summary:")
            print(f"  Header: {results['extraction_status']['header']}")
            print(f"  Items: {results['extraction_status']['items']}")
            print(f"  Summary: {results['extraction_status']['summary']}")
            print(f"  Overall: {results['extraction_status']['overall']}")

            # Return the results
            return results

        except Exception as e:
            handle_exception(
                func_name="extract_invoice_tables",
                exception=e,
                context={"pdf_path": pdf_path}
            )
            if "pdf_document" in locals():
                pdf_document.close()
            return None



    def _process_extracted_table(self, table_df, section, results, region_index=0, page_number=1):
        """Process an extracted table and add it to results

        Args:
            table_df: Extracted DataFrame
            section: Section name ('header', 'items', 'summary')
            results: Results dictionary to update
            region_index: Index of the region (0-based)
            page_number: Page number (1-based)
        """
        try:
            # Basic cleaning - only replace empty strings with NaN
            table_df = table_df.replace(r"^\s*$", pd.NA, regex=True)

            # Standardize region labels for consistency
            table_df = standardize_dataframe_labels(
                df=table_df,
                section=section,
                region_index=region_index,
                page_number=page_number
            )

            log_info(f"Using raw extraction data without regex patterns for {section}")

            # Store the table
            if not table_df.empty:
                if section == "header":
                    results["header_tables"].append(table_df)
                    log_info(f"Extracted header table with {len(table_df)} rows")
                    results["extraction_status"]["header"] = "success"
                elif section == "items":
                    results["items_tables"].append(table_df)
                    log_info(f"Extracted items table with {len(table_df)} rows")
                    results["extraction_status"]["items"] = "success"
                else:  # summary
                    results["summary_tables"].append(table_df)
                    log_info(f"Extracted summary table with {len(table_df)} rows")
                    results["extraction_status"]["summary"] = "success"
            else:
                log_info(f"Table is empty after processing for {section}")

        except Exception as e:
            handle_exception(
                func_name="_process_extracted_table",
                exception=e,
                context={"section": section, "table_shape": str(table_df.shape) if table_df is not None else "None"}
            )

    def stop_processing(self):
        """Stop the processing of files"""
        self.should_stop = True
        self.status_label.setText("Stopping processing...")

    def update_processing_time(self, is_final=False):
        """Update the processing time display"""
        if self.start_time:
            elapsed_time = time.time() - self.start_time
            minutes, seconds = divmod(int(elapsed_time), 60)
            hours, minutes = divmod(minutes, 60)

            if hours > 0:
                time_str = f"Time: {hours}h {minutes}m {seconds}s"
            elif minutes > 0:
                time_str = f"Time: {minutes}m {seconds}s"
            else:
                time_str = f"Time: {seconds}s"

            if is_final:
                time_str = f"Total {time_str}"

            self.processing_time_label.setText(time_str)

    def get_selected_template_id(self):
        """Get the ID of the selected template"""
        if self.template_combo.count() == 0:
            return None
        return self.template_combo.currentData()

    def process_with_invoice2data(self, pdf_path, template_data, extracted_data=None):
        """Process a PDF file with invoice2data using the template and extracted data

        Args:
            pdf_path (str): Path to the PDF file
            template_data (dict): Template data from the database
            extracted_data (dict, optional): Extracted data from the PDF. If None, gets the latest data from processed_data.

        Returns:
            dict: The extraction result from invoice2data, or None if extraction failed
        """
        if not INVOICE2DATA_AVAILABLE:
            print("invoice2data module is not available. Skipping invoice2data processing.")
            return None

        try:
            # Use extracted_data if provided, otherwise get from processed_data
            if extracted_data is None:
                extracted_data = self.processed_data.get(pdf_path, {})
                print(f"Using extracted data from processed_data")

            # Use the unified invoice processing utilities
            print(f"Using invoice_processing_utils.process_with_invoice2data")
            result = invoice_processing_utils.process_with_invoice2data(
                pdf_path=pdf_path,
                template_data=template_data,
                extracted_data=extracted_data,
                temp_dir=self.temp_dir
            )

            return result

        except Exception as e:
            error_message = str(e)
            print(f"Error in process_with_invoice2data: {error_message}")
            import traceback
            traceback.print_exc()
            return None

    def export_invoice2data_results(self, invoice2data_results):
        """Export invoice2data results to a separate JSON file"""
        if not invoice2data_results:
            print("No invoice2data results to export")
            return None

        try:
            # Create export directory if it doesn't exist
            export_dir = os.path.abspath("exported_data")
            os.makedirs(export_dir, exist_ok=True)

            # Save to JSON file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(export_dir, f"invoice2data_results_{timestamp}.json")

            print(f"Attempting to save file to: {filename}")
            print(f"Results to export: {len(invoice2data_results)} items")

            # Debug the results
            for pdf_name, result in invoice2data_results.items():
                print(f"Result for {pdf_name}: {type(result)}")
                if result is None:
                    print(f"Warning: Result for {pdf_name} is None")
                elif not isinstance(result, dict):
                    print(f"Warning: Result for {pdf_name} is not a dictionary: {type(result)}")
                    print(f"Content: {result}")

            # Use invoice2data's built-in JSON serialization
            # This avoids issues with datetime serialization
            with open(filename, "w", encoding="utf-8") as f:
                # Convert the results to a list format that invoice2data expects
                results_list = []
                for pdf_name, result in invoice2data_results.items():
                    if result is None:
                        # Skip None results
                        continue

                    if isinstance(result, dict):
                        # Add filename to the result
                        result_copy = result.copy()
                        result_copy['filename'] = pdf_name
                        results_list.append(result_copy)
                    elif isinstance(result, str):
                        # Try to parse JSON string
                        try:
                            result_dict = json.loads(result)
                            if isinstance(result_dict, dict):
                                result_dict['filename'] = pdf_name
                                results_list.append(result_dict)
                            elif isinstance(result_dict, list) and len(result_dict) > 0:
                                # If it's a list, add each item
                                for item in result_dict:
                                    if isinstance(item, dict):
                                        item_copy = item.copy()
                                        item_copy['filename'] = pdf_name
                                        results_list.append(item_copy)
                        except json.JSONDecodeError:
                            print(f"Warning: Could not parse result for {pdf_name} as JSON")
                    elif isinstance(result, list):
                        # If it's already a list, add each item
                        for item in result:
                            if isinstance(item, dict):
                                item_copy = item.copy()
                                item_copy['filename'] = pdf_name
                                results_list.append(item_copy)

                # If results_list is empty, create a minimal valid JSON array
                if not results_list:
                    print("Warning: No valid results to export, creating empty array")
                    results_list = []

                # Use simple JSON serialization with str conversion for non-serializable types
                json.dump(results_list, f, indent=2, ensure_ascii=False, default=str)
                print(f"Exported {len(results_list)} results to {filename}")

            # Verify the file was created successfully
            if os.path.exists(filename):
                print(f"Exported invoice2data results to {filename}")
                return os.path.abspath(filename)  # Return absolute path
            else:
                print(f"Error: Failed to create file {filename}")
                return None
        except Exception as e:
            print(f"Error exporting invoice2data results: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    def open_validation_screen(self):
        """Process selected PDFs with the selected template using invoice2data and update validation status"""
        if not self.processed_data:
            QMessageBox.warning(self, "Warning", "No data to validate. Please process files first.")
            return

        # Get template ID for invoice2data processing
        template_id = self.get_selected_template_id()
        if not template_id:
            QMessageBox.warning(self, "Warning", "Please select a template first.")
            return

        # Update status
        self.status_label.setText("Validating with invoice2data...")
        self.status_label.setStyleSheet(f"""
            padding: 4px 8px;
            border-radius: 4px;
            background-color: {self.theme['primary'] + '20'};
            color: {self.theme['primary']};
            font-weight: bold;
        """)
        QApplication.processEvents()  # Ensure UI updates

        # Setup temporary directory for storing txt files and templates
        self.setup_temp_directory()

        # Get the template data from the database
        template_data = self.fetch_template_from_database(template_id)

        # Check if we have a valid template
        if not template_data:
            UIMessageFactory.show_warning(self, "Template Error", "No valid invoice2data template found. Please create a JSON template first.")
            self.status_label.setText("No valid invoice2data template found")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['warning'] + '20'};
                color: {self.theme['warning']};
                font-weight: bold;
            """)
            self.cleanup_temp_directory()
            return

        # Get the list of PDF paths to process
        pdf_paths = list(self.processed_data.keys())

        # Process the PDFs with the selected template
        validation_results = {}
        for row in range(self.results_table.rowCount()):
            file_name = self.results_table.item(row, 0).text()

            # Find the corresponding PDF path
            pdf_path = None
            for path in pdf_paths:
                if os.path.basename(path) == file_name:
                    pdf_path = path
                    break

            if pdf_path and pdf_path in self.processed_data:
                # Update validation status to "Processing..."
                validation_status_item = QTableWidgetItem("Processing...")
                validation_status_item.setData(Qt.UserRole, "processing")
                validation_status_item.setForeground(QColor("#3B82F6"))  # Blue color for processing
                self.results_table.setItem(row, 2, validation_status_item)
                QApplication.processEvents()  # Keep UI responsive

                # Process the PDF with the selected template
                success, result, used_template_id, warnings = self.process_pdf_with_selected_template(pdf_path, template_id)

                # Update validation status based on result
                if success:
                    validation_status_item = QTableWidgetItem("Valid")
                    validation_status_item.setData(Qt.UserRole, "valid")
                    validation_status_item.setForeground(QColor("#10B981"))  # Green color for valid
                    validation_results[pdf_path] = result
                else:
                    validation_status_item = QTableWidgetItem("Invalid")
                    validation_status_item.setData(Qt.UserRole, "invalid")
                    validation_status_item.setForeground(QColor("#EF4444"))  # Red color for invalid

                self.results_table.setItem(row, 2, validation_status_item)
                QApplication.processEvents()  # Keep UI responsive
        # Check if we have any successful results
        successful_results = {k: v for k, v in validation_results.items() if v is not None}

        if successful_results:
            # Convert the results to the format expected by export_invoice2data_results
            invoice2data_results = {}
            for pdf_path, result in successful_results.items():
                pdf_filename = os.path.basename(pdf_path)
                invoice2data_results[pdf_filename] = result

            # Export the results (default to JSON format)
            result_file = self.export_invoice2data_results(invoice2data_results, export_format="json")

            if result_file:
                self.status_label.setText("Validation complete")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['secondary'] + '20'};
                    color: {self.theme['secondary']};
                    font-weight: bold;
                """)

                # Show success message
                success_box = QMessageBox(self)
                success_box.setWindowTitle("Validation Complete")
                success_box.setIcon(QMessageBox.Information)
                success_box.setText("Validation completed successfully")
                success_box.setInformativeText(
                    f"<b>Results exported to:</b><br>{result_file}<br><br>"
                    f"<b>Files processed:</b> {len(invoice2data_results)}<br><br>"
                    f"<b>Successful:</b> {len(successful_results)} of {len(pdf_paths)}"
                )

                # Create export directory if it doesn't exist
                export_dir = os.path.dirname(result_file)
                os.makedirs(export_dir, exist_ok=True)

                # Verify the file exists before trying to open it
                if os.path.exists(result_file):
                    print(f"File exists and will be opened: {result_file}")

                    # Open folder button
                    open_folder_btn = success_box.addButton("Open Folder", QMessageBox.ActionRole)
                    def open_folder():
                        try:
                            folder_path = os.path.dirname(result_file)
                            print(f"Opening folder: {folder_path}")
                            os.startfile(folder_path)
                        except Exception as e:
                            print(f"Error opening folder: {str(e)}")
                            QMessageBox.warning(self, "Error", f"Could not open folder: {str(e)}")

                    open_folder_btn.clicked.connect(open_folder)

                    # Open file button
                    open_file_btn = success_box.addButton("Open File", QMessageBox.ActionRole)
                    def open_file():
                        try:
                            print(f"Opening file: {result_file}")
                            os.startfile(result_file)
                        except Exception as e:
                            print(f"Error opening file: {str(e)}")
                            QMessageBox.warning(self, "Error", f"Could not open file: {str(e)}")

                    open_file_btn.clicked.connect(open_file)

                    # Export as JSON button (if current file is not already JSON)
                    if not result_file.lower().endswith('.json'):
                        export_json_btn = success_box.addButton("Export as JSON", QMessageBox.ActionRole)
                        def export_as_json():
                            try:
                                # Generate a new JSON filename
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                json_file = os.path.join(export_dir, f"invoice2data_results_{timestamp}.json")

                                # Export to JSON
                                json_result_file = self.export_invoice2data_results(invoice2data_results, export_format="json", custom_filename=json_file)

                                if json_result_file and os.path.exists(json_result_file):
                                    # Show success message
                                    QMessageBox.information(
                                        self,
                                        "JSON Export Successful",
                                        f"Results exported to JSON:\n{json_result_file}"
                                    )

                                    # Open the file
                                    try:
                                        os.startfile(json_result_file)
                                    except Exception as e:
                                        print(f"Error opening exported JSON file: {str(e)}")
                                else:
                                    QMessageBox.warning(self, "Export Error", "Failed to export as JSON")
                            except Exception as e:
                                print(f"Error exporting as JSON: {str(e)}")
                                QMessageBox.warning(self, "Export Error", f"Could not export as JSON: {str(e)}")

                        export_json_btn.clicked.connect(export_as_json)

                    # Export as CSV button (if current file is not already CSV)
                    if not result_file.lower().endswith('.csv'):
                        export_csv_btn = success_box.addButton("Export as CSV", QMessageBox.ActionRole)
                        def export_as_csv():
                            try:
                                # Generate a new CSV filename
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                csv_file = os.path.join(export_dir, f"invoice2data_results_{timestamp}.csv")

                                # Export to CSV
                                csv_result_file = self.export_invoice2data_results(invoice2data_results, export_format="csv", custom_filename=csv_file)

                                if csv_result_file and os.path.exists(csv_result_file):
                                    # Show success message
                                    QMessageBox.information(
                                        self,
                                        "CSV Export Successful",
                                        f"Results exported to CSV:\n{csv_result_file}"
                                    )

                                    # Open the file
                                    try:
                                        os.startfile(csv_result_file)
                                    except Exception as e:
                                        print(f"Error opening exported CSV file: {str(e)}")
                                else:
                                    QMessageBox.warning(self, "Export Error", "Failed to export as CSV")
                            except Exception as e:
                                print(f"Error exporting as CSV: {str(e)}")
                                QMessageBox.warning(self, "Export Error", f"Could not export as CSV: {str(e)}")

                        export_csv_btn.clicked.connect(export_as_csv)

                    # Clean temp files button
                    clean_temp_btn = success_box.addButton("Clean Temp Files", QMessageBox.ActionRole)
                    clean_temp_btn.clicked.connect(self.cleanup_temp_directory)
                else:
                    # File doesn't exist, show a warning
                    print(f"Warning: Result file {result_file} does not exist")
                    success_box.setInformativeText(
                        f"<b>Warning:</b> The result file could not be found.<br><br>"
                        f"<b>Expected file:</b> {result_file}<br><br>"
                        f"<b>Files processed:</b> {len(invoice2data_results)}"
                    )

                # OK button
                ok_btn = success_box.addButton(QMessageBox.Ok)
                ok_btn.setDefault(True)

                success_box.exec()
            else:
                self.status_label.setText("Error exporting validation results")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['danger'] + '20'};
                    color: {self.theme['danger']};
                    font-weight: bold;
                """)
                UIMessageFactory.show_error(self, "Export Error", "Failed to export validation results")
                self.cleanup_temp_directory()
        else:
            self.status_label.setText("No valid results to export")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['warning'] + '20'};
                color: {self.theme['warning']};
                font-weight: bold;
            """)
            UIMessageFactory.show_warning(self, "No Valid Results", "No valid results were generated. Check the template and try again.")
            self.cleanup_temp_directory()
    def process_pdf_with_selected_template(self, pdf_path, template_id=None):
        """Process a PDF with the selected template from the dropdown

        Args:
            pdf_path (str): Path to the PDF file to process
            template_id (int, optional): ID of the template to use. If None, uses the selected template from dropdown.

        Returns:
            tuple: (success, result, template_id, warnings) where:
                - success (bool): True if extraction was successful
                - result (dict): The extracted data or None if extraction failed
                - template_id (int): ID of the template that was used
                - warnings (str): Any warnings generated during extraction
        """
        # If no template_id provided, get the selected template from the dropdown
        if template_id is None:
            template_id = self.get_selected_template_id()

        if not template_id:
            return False, None, None, "No template selected. Please select a template from the dropdown."

        # Create a temporary directory for our files
        if not self.temp_dir:
            self.setup_temp_directory()

        # Fetch template data from database
        template_data = self.fetch_template_from_database(template_id)
        if not template_data:
            return False, None, None, f"Template with ID {template_id} not found or has no JSON template"
        # Get template name
        try:
            conn = sqlite3.connect("invoice_templates.db")
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM templates WHERE id = ?", (template_id,))
            template_name = cursor.fetchone()[0]
            conn.close()
        except Exception as e:
            print(f"Error fetching template name: {str(e)}")
            template_name = f"template_{template_id}"

        # Get the extracted data from the PDF
        extracted_data = self.processed_data.get(pdf_path, {})

        # Process with invoice2data
        result = self.process_with_invoice2data(pdf_path, template_data, extracted_data)

        # Set up logging to capture warnings
        import io
        import logging
        invoice2data_logger = logging.getLogger("invoice2data")
        log_capture = io.StringIO()
        string_handler = logging.StreamHandler(log_capture)
        string_handler.setLevel(logging.WARNING)
        invoice2data_logger.addHandler(string_handler)
        warnings = log_capture.getvalue().strip()

        # Check if we got a result
        if result:
            return True, result, template_id, warnings
        else:
            return False, None, template_id, f"Failed to extract data with invoice2data. {warnings}"






    def export_invoice2data_results(self, invoice2data_results, export_format="json", custom_filename=None):
        """Export invoice2data results to a file in the specified format

        Args:
            invoice2data_results (dict): Dictionary of invoice2data results
            export_format (str): Format to export as ("json" or "csv")
            custom_filename (str, optional): Custom filename to use. If None, generates a timestamped filename.

        Returns:
            str: Path to the exported file, or None if export failed
        """
        if not invoice2data_results:
            print("No invoice2data results to export")
            return None

        try:
            # Create export directory if it doesn't exist
            export_dir = os.path.abspath("exported_data")
            os.makedirs(export_dir, exist_ok=True)

            # Generate filename if not provided
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if custom_filename:
                # Use the provided filename but ensure it has the correct extension
                base_name = os.path.splitext(custom_filename)[0]
                filename = f"{base_name}.{export_format.lower()}"
                filename = os.path.join(export_dir, filename)
            else:
                # Generate a timestamped filename
                filename = os.path.join(export_dir, f"invoice2data_results_{timestamp}.{export_format.lower()}")

            print(f"Attempting to save file to: {filename}")
            print(f"Results to export: {len(invoice2data_results)} items")
            print(f"Export format: {export_format}")

            # Debug the results
            for pdf_name, result in invoice2data_results.items():
                print(f"Result for {pdf_name}: {type(result)}")
                if result is None:
                    print(f"Warning: Result for {pdf_name} is None")
                elif not isinstance(result, dict):
                    print(f"Warning: Result for {pdf_name} is not a dictionary: {type(result)}")
                    print(f"Content: {result}")

            # Convert the results to a list format
            results_list = []
            for pdf_name, result in invoice2data_results.items():
                if result is None:
                    # Skip None results
                    continue

                if isinstance(result, dict):
                    # Add filename to the result
                    result_copy = result.copy()
                    result_copy['filename'] = pdf_name
                    results_list.append(result_copy)
                elif isinstance(result, str):
                    # Try to parse JSON string
                    try:
                        result_dict = json.loads(result)
                        if isinstance(result_dict, dict):
                            result_dict['filename'] = pdf_name
                            results_list.append(result_dict)
                        elif isinstance(result_dict, list) and len(result_dict) > 0:
                            # If it's a list, add each item
                            for item in result_dict:
                                if isinstance(item, dict):
                                    item_copy = item.copy()
                                    item_copy['filename'] = pdf_name
                                    results_list.append(item_copy)
                    except json.JSONDecodeError:
                        print(f"Warning: Could not parse result for {pdf_name} as JSON")
                elif isinstance(result, list):
                    # If it's already a list, add each item
                    for item in result:
                        if isinstance(item, dict):
                            item_copy = item.copy()
                            item_copy['filename'] = pdf_name
                            results_list.append(item_copy)

            # If results_list is empty, create a minimal valid array
            if not results_list:
                print("Warning: No valid results to export, creating empty array")
                results_list = []

            # Export based on the requested format
            if export_format.lower() == "json":
                # Export as JSON
                with open(filename, "w", encoding="utf-8") as f:
                    # Use simple JSON serialization with str conversion for non-serializable types
                    json.dump(results_list, f, indent=2, ensure_ascii=False, default=str)
                    print(f"Exported {len(results_list)} results to JSON file: {filename}")

            elif export_format.lower() == "csv":
                # Export as CSV
                import csv

                # Get all unique keys from all results to use as CSV headers
                all_keys = set()
                for result in results_list:
                    # Handle nested line items by flattening them
                    flat_result = self._flatten_nested_items(result)
                    all_keys.update(flat_result.keys())

                # Sort keys to ensure consistent column order
                # Put common important fields first
                priority_keys = ['filename', 'invoice_number', 'date', 'amount', 'issuer']
                sorted_keys = sorted(all_keys, key=lambda x: (
                    0 if x in priority_keys else 1,  # Priority keys first
                    priority_keys.index(x) if x in priority_keys else 999,  # Order of priority keys
                    x  # Alphabetical for the rest
                ))

                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=sorted_keys)
                    writer.writeheader()

                    for result in results_list:
                        # Flatten nested items before writing to CSV
                        flat_result = self._flatten_nested_items(result)

                        # Convert all values to strings to avoid type issues
                        row_dict = {}
                        for key, value in flat_result.items():
                            if isinstance(value, (datetime, date)):
                                row_dict[key] = value.isoformat()
                            elif value is None:
                                row_dict[key] = ''
                            else:
                                row_dict[key] = str(value)

                        writer.writerow(row_dict)

                    print(f"Exported {len(results_list)} results to CSV file: {filename}")
            else:
                print(f"Unsupported export format: {export_format}")
                return None

            # Verify the file was created successfully
            if os.path.exists(filename):
                print(f"Exported invoice2data results to {filename}")
                return os.path.abspath(filename)  # Return absolute path
            else:
                print(f"Error: Failed to create file {filename}")
                return None
        except Exception as e:
            print(f"Error exporting invoice2data results: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def _flatten_nested_items(self, result_dict):
        """Flatten nested line items in invoice2data results for CSV export

        Args:
            result_dict (dict): Dictionary containing invoice2data results

        Returns:
            dict: Flattened dictionary with nested items converted to columns
        """
        if not isinstance(result_dict, dict):
            return result_dict

        flat_dict = result_dict.copy()

        # Check for common line items fields
        for field in ['lines', 'line_items', 'items']:
            if field in flat_dict and isinstance(flat_dict[field], list) and len(flat_dict[field]) > 0:
                items = flat_dict.pop(field)

                # Process each line item
                for i, item in enumerate(items):
                    if isinstance(item, dict):
                        # Add each field from the line item with a prefix
                        for key, value in item.items():
                            flat_dict[f"{field}_{i+1}_{key}"] = value
                    else:
                        # If it's not a dict, just add the whole item
                        flat_dict[f"{field}_{i+1}"] = item

        return flat_dict

    def open_validation_screen(self):
        """Process selected PDFs with the selected template using invoice2data and update validation status"""
        if not self.processed_data:
            QMessageBox.warning(self, "Warning", "No data to validate. Please process files first.")
            return

        # Get template ID for invoice2data processing
        template_id = self.get_selected_template_id()
        if not template_id:
            QMessageBox.warning(self, "Warning", "Please select a template first.")
            return

        # Update status
        self.status_label.setText("Validating with invoice2data...")
        self.status_label.setStyleSheet(f"""
            padding: 4px 8px;
            border-radius: 4px;
            background-color: {self.theme['primary'] + '20'};
            color: {self.theme['primary']};
            font-weight: bold;
        """)
        QApplication.processEvents()  # Ensure UI updates

        # Setup temporary directory for storing txt files and templates
        self.setup_temp_directory()

        # Get the template data from the database
        template_data = self.fetch_template_from_database(template_id)

        # Check if we have a valid template
        if not template_data:
            QMessageBox.warning(self, "Template Error", "No valid invoice2data template found. Please create a JSON template first.")
            self.status_label.setText("No valid invoice2data template found")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['warning'] + '20'};
                color: {self.theme['warning']};
                font-weight: bold;
            """)
            self.cleanup_temp_directory()
            return

        # Get the list of PDF paths to process
        pdf_paths = list(self.processed_data.keys())

        # Process the PDFs with the selected template
        validation_results = {}
        for row in range(self.results_table.rowCount()):
            file_name = self.results_table.item(row, 0).text()

            # Find the corresponding PDF path
            pdf_path = None
            for path in pdf_paths:
                if os.path.basename(path) == file_name:
                    pdf_path = path
                    break

            if pdf_path and pdf_path in self.processed_data:
                # Update validation status to "Processing..."
                validation_status_item = QTableWidgetItem("Processing...")
                validation_status_item.setData(Qt.UserRole, "processing")
                validation_status_item.setForeground(QColor("#3B82F6"))  # Blue color for processing
                self.results_table.setItem(row, 2, validation_status_item)
                QApplication.processEvents()  # Keep UI responsive

                # Process the PDF with the selected template
                success, result, used_template_id, warnings = self.process_pdf_with_selected_template(pdf_path, template_id)

                # Update validation status based on result
                if success:
                    validation_status_item = QTableWidgetItem("Valid")
                    validation_status_item.setData(Qt.UserRole, "valid")
                    validation_status_item.setForeground(QColor("#10B981"))  # Green color for valid
                    validation_results[pdf_path] = result
                else:
                    validation_status_item = QTableWidgetItem("Invalid")
                    validation_status_item.setData(Qt.UserRole, "invalid")
                    validation_status_item.setForeground(QColor("#EF4444"))  # Red color for invalid

                self.results_table.setItem(row, 2, validation_status_item)
                QApplication.processEvents()  # Keep UI responsive

        # Check if we have any successful results
        successful_results = {k: v for k, v in validation_results.items() if v is not None}

        if successful_results:
            # Convert the results to the format expected by export_invoice2data_results
            invoice2data_results = {}
            for pdf_path, result in successful_results.items():
                pdf_filename = os.path.basename(pdf_path)
                invoice2data_results[pdf_filename] = result

            # Export the results (default to JSON format)
            result_file = self.export_invoice2data_results(invoice2data_results, export_format="json")

            if result_file:
                self.status_label.setText("Validation complete")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['secondary'] + '20'};
                    color: {self.theme['secondary']};
                    font-weight: bold;
                """)

                # Show success message
                success_box = QMessageBox(self)
                success_box.setWindowTitle("Validation Complete")
                success_box.setIcon(QMessageBox.Information)
                success_box.setText("Validation completed successfully")
                success_box.setInformativeText(
                    f"<b>Results exported to:</b><br>{result_file}<br><br>"
                    f"<b>Files processed:</b> {len(invoice2data_results)}<br><br>"
                    f"<b>Successful:</b> {len(successful_results)} of {len(pdf_paths)}"
                )

                # Create export directory if it doesn't exist
                export_dir = os.path.dirname(result_file)
                os.makedirs(export_dir, exist_ok=True)

                # Verify the file exists before trying to open it
                if os.path.exists(result_file):
                    print(f"File exists and will be opened: {result_file}")

                    # Open folder button
                    open_folder_btn = success_box.addButton("Open Folder", QMessageBox.ActionRole)
                    def open_folder():
                        try:
                            folder_path = os.path.dirname(result_file)
                            print(f"Opening folder: {folder_path}")
                            os.startfile(folder_path)
                        except Exception as e:
                            print(f"Error opening folder: {str(e)}")
                            QMessageBox.warning(self, "Error", f"Could not open folder: {str(e)}")

                    open_folder_btn.clicked.connect(open_folder)

                    # Open file button
                    open_file_btn = success_box.addButton("Open File", QMessageBox.ActionRole)
                    def open_file():
                        try:
                            print(f"Opening file: {result_file}")
                            os.startfile(result_file)
                        except Exception as e:
                            print(f"Error opening file: {str(e)}")
                            QMessageBox.warning(self, "Error", f"Could not open file: {str(e)}")

                    open_file_btn.clicked.connect(open_file)

                    # Export as JSON button (if current file is not already JSON)
                    if not result_file.lower().endswith('.json'):
                        export_json_btn = success_box.addButton("Export as JSON", QMessageBox.ActionRole)
                        def export_as_json():
                            try:
                                # Generate a new JSON filename
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                json_file = os.path.join(export_dir, f"invoice2data_results_{timestamp}.json")

                                # Export to JSON
                                json_result_file = self.export_invoice2data_results(invoice2data_results, export_format="json", custom_filename=json_file)

                                if json_result_file and os.path.exists(json_result_file):
                                    # Show success message
                                    QMessageBox.information(
                                        self,
                                        "JSON Export Successful",
                                        f"Results exported to JSON:\n{json_result_file}"
                                    )

                                    # Open the file
                                    try:
                                        os.startfile(json_result_file)
                                    except Exception as e:
                                        print(f"Error opening exported JSON file: {str(e)}")
                                else:
                                    QMessageBox.warning(self, "Export Error", "Failed to export as JSON")
                            except Exception as e:
                                print(f"Error exporting as JSON: {str(e)}")
                                QMessageBox.warning(self, "Export Error", f"Could not export as JSON: {str(e)}")

                        export_json_btn.clicked.connect(export_as_json)

                    # Export as CSV button (if current file is not already CSV)
                    if not result_file.lower().endswith('.csv'):
                        export_csv_btn = success_box.addButton("Export as CSV", QMessageBox.ActionRole)
                        def export_as_csv():
                            try:
                                # Generate a new CSV filename
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                csv_file = os.path.join(export_dir, f"invoice2data_results_{timestamp}.csv")

                                # Export to CSV
                                csv_result_file = self.export_invoice2data_results(invoice2data_results, export_format="csv", custom_filename=csv_file)

                                if csv_result_file and os.path.exists(csv_result_file):
                                    # Show success message
                                    QMessageBox.information(
                                        self,
                                        "CSV Export Successful",
                                        f"Results exported to CSV:\n{csv_result_file}"
                                    )

                                    # Open the file
                                    try:
                                        os.startfile(csv_result_file)
                                    except Exception as e:
                                        print(f"Error opening exported CSV file: {str(e)}")
                                else:
                                    QMessageBox.warning(self, "Export Error", "Failed to export as CSV")
                            except Exception as e:
                                print(f"Error exporting as CSV: {str(e)}")
                                QMessageBox.warning(self, "Export Error", f"Could not export as CSV: {str(e)}")

                        export_csv_btn.clicked.connect(export_as_csv)

                    # Clean temp files button
                    clean_temp_btn = success_box.addButton("Clean Temp Files", QMessageBox.ActionRole)
                    clean_temp_btn.clicked.connect(self.cleanup_temp_directory)
                else:
                    # File doesn't exist, show a warning
                    print(f"Warning: Result file {result_file} does not exist")
                    success_box.setInformativeText(
                        f"<b>Warning:</b> The result file could not be found.<br><br>"
                        f"<b>Expected file:</b> {result_file}<br><br>"
                        f"<b>Files processed:</b> {len(invoice2data_results)}"
                    )

                # OK button
                ok_btn = success_box.addButton(QMessageBox.Ok)
                ok_btn.setDefault(True)

                success_box.exec()
            else:
                self.status_label.setText("Error exporting validation results")
                self.status_label.setStyleSheet(f"""
                    padding: 4px 8px;
                    border-radius: 4px;
                    background-color: {self.theme['danger'] + '20'};
                    color: {self.theme['danger']};
                    font-weight: bold;
                """)
                QMessageBox.critical(self, "Export Error", "Failed to export validation results")
                self.cleanup_temp_directory()
        else:
            self.status_label.setText("No valid results to export")
            self.status_label.setStyleSheet(f"""
                padding: 4px 8px;
                border-radius: 4px;
                background-color: {self.theme['warning'] + '20'};
                color: {self.theme['warning']};
                font-weight: bold;
            """)
            QMessageBox.warning(self, "No Valid Results", "No valid results were generated. Check the template and try again.")
            self.cleanup_temp_directory()









    def process_pdf_with_selected_template(self, pdf_path, template_id=None):
        """Process a PDF with the selected template from the dropdown

        Args:
            pdf_path (str): Path to the PDF file to process
            template_id (int, optional): ID of the template to use. If None, uses the selected template from dropdown.

        Returns:
            tuple: (success, result, template_id, warnings) where:
                - success (bool): True if extraction was successful
                - result (dict): The extracted data or None if extraction failed
                - template_id (int): ID of the template that was used
                - warnings (str): Any warnings generated during extraction
        """
        # If no template_id provided, get the selected template from the dropdown
        if template_id is None:
            template_id = self.get_selected_template_id()

        if not template_id:
            return False, None, None, "No template selected. Please select a template from the dropdown."

        # Create a temporary directory for our files
        if not self.temp_dir:
            self.setup_temp_directory()

        # Fetch template data from database
        template_data = self.fetch_template_from_database(template_id)
        if not template_data:
            return False, None, None, f"Template with ID {template_id} not found or has no JSON template"

        # Get template name
        try:
            conn = sqlite3.connect("invoice_templates.db")
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM templates WHERE id = ?", (template_id,))
            template_name = cursor.fetchone()[0]
            conn.close()
        except Exception as e:
            print(f"Error fetching template name: {str(e)}")
            template_name = f"template_{template_id}"

        # Get the extracted data from the PDF
        extracted_data = self.processed_data.get(pdf_path, {})

        # Convert the extracted data to text
        pdf_text = self.convert_extraction_to_text(pdf_path, extracted_data)

        # Get a base filename from the PDF
        base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
        # Clean the filename to remove invalid characters
        base_filename = re.sub(r'[^\w\-_.]', '_', base_filename)

        # Create a common name for both text file and template
        common_name = f"temp_{base_filename}"

        # Create temporary text file
        temp_text_path = os.path.join(self.temp_dir, f"{common_name}.txt")
        with open(temp_text_path, "w", encoding="utf-8") as f:
            f.write(pdf_text)
        print(f"Saved extracted text to {temp_text_path}")

        # Create a temporary directory for templates
        templates_dir = os.path.join(self.temp_dir, "templates")
        os.makedirs(templates_dir, exist_ok=True)

        # Use the same common name for the template
        safe_template_name = common_name

        # Debug the template data
        print(f"[DEBUG] Template data type: {type(template_data)}")
        print(f"[DEBUG] Template data keys: {template_data.keys() if isinstance(template_data, dict) else 'Not a dictionary'}")

        # Extract just the JSON template part
        json_template = template_data.get("json_template", {})
        print(f"[DEBUG] JSON template type: {type(json_template)}")
        print(f"[DEBUG] JSON template keys: {json_template.keys() if isinstance(json_template, dict) else 'Not a dictionary'}")

        if not json_template:
            print(f"[DEBUG] Template has no JSON template data. Creating a basic template using factory.")
            # Create a basic template using factory
            base_template = TemplateFactory.create_default_invoice_template(template_name, "INR")
            json_template = {
                "issuer": base_template["issuer"],
                "fields": {
                    "invoice_number": {
                        "parser": "regex",
                        "regex": "invoice_number:\\s*([^\\n]+)"
                    },
                    "date": {
                        "parser": "regex",
                        "regex": "date:\\s*([^\\n]+)"
                    },
                    "amount": {
                        "parser": "regex",
                        "regex": "amount:\\s*([^\\n]+)"
                    }
                },
                "keywords": [safe_template_name],
                "options": base_template["options"]
            }
        else:
            print(f"[DEBUG] Found JSON template data")

        # Ensure the template has the required fields
        if "fields" not in json_template:
            print(f"[DEBUG] Adding missing 'fields' section to template")
            json_template["fields"] = {}

        # Make sure the template has the three mandatory fields: amount, date, and invoice_number
        for field in ["amount", "date", "invoice_number"]:
            if field not in json_template["fields"]:
                print(f"[DEBUG] Adding missing field '{field}' to template")
                json_template["fields"][field] = {
                    "parser": "regex",
                    "regex": f"{field}:\\s*([^\\n]+)"
                }

        # Add keywords if missing
        if "keywords" not in json_template or not json_template["keywords"]:
            print(f"[DEBUG] Adding missing 'keywords' to template")
            # Use the template name as a keyword
            json_template["keywords"] = [safe_template_name]

        # Add issuer if missing
        if "issuer" not in json_template:
            print(f"[DEBUG] Adding missing 'issuer' to template")
            json_template["issuer"] = safe_template_name

        # Add required fields for invoice2data
        if "name" not in json_template:
            json_template["name"] = safe_template_name
        if "template_name" not in json_template:
            json_template["template_name"] = safe_template_name

        # Save the template as YAML (preferred format)
        try:
            import yaml
            template_path = os.path.join(templates_dir, f"{safe_template_name}.yml")
            with open(template_path, "w", encoding="utf-8") as f:
                yaml.dump(json_template, f, default_flow_style=False, allow_unicode=True)
            print(f"Saved template {template_id} to {template_path} (YAML)")
        except ImportError:
            print("YAML module not available, using JSON instead")
            # Fallback to JSON if YAML is not available
            template_path = os.path.join(templates_dir, f"{safe_template_name}.json")
            with open(template_path, "w", encoding="utf-8") as f:
                json.dump(json_template, f, indent=2, ensure_ascii=False)
            print(f"Saved template {template_id} to {template_path} (JSON)")

        # Define output name (without extension)
        output_name = f"{base_filename}_result"

        # Set up logging to capture warnings
        import io
        import logging
        invoice2data_logger = logging.getLogger("invoice2data")
        log_capture = io.StringIO()
        string_handler = logging.StreamHandler(log_capture)
        string_handler.setLevel(logging.WARNING)
        invoice2data_logger.addHandler(string_handler)

        # Build the command
        cmd = [
            "invoice2data",
            "--input-reader", "text",
            "--exclude-built-in-templates",
            "--template-folder", templates_dir,
            "--output-format", "json",
            "--output-name", os.path.join(self.temp_dir, output_name),
            "--debug",
            temp_text_path
        ]

        print(f"[DEBUG] Command: {' '.join(cmd)}")
        print(f"[DEBUG] Using template folder: {templates_dir}")
        print(f"[DEBUG] Using text file: {temp_text_path}")
        print(f"[DEBUG] Output name: {os.path.join(self.temp_dir, output_name)}")

        # Print the command for debugging
        print(f"[DEBUG] Running command: {' '.join(cmd)}")

        # Print the content of the text file
        try:
            with open(temp_text_path, 'r', encoding='utf-8') as f:
                text_content = f.read()
                print(f"[DEBUG] Text file content (first 500 chars): {text_content[:500]}...")
        except Exception as e:
            print(f"[DEBUG] Error reading text file: {str(e)}")

        # Print the content of the template file
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template_content = f.read()
                print(f"[DEBUG] Template file content: {template_content}")
        except Exception as e:
            print(f"[DEBUG] Error reading template file: {str(e)}")

        # Run the command and capture output
        import subprocess
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            check=False
        )

        # Print the command output
        print(f"[DEBUG] Command return code: {result.returncode}")
        print(f"[DEBUG] Command stdout: {result.stdout}")
        print(f"[DEBUG] Command stderr: {result.stderr}")

        # Get warnings
        warnings = log_capture.getvalue().strip()

        # Check if command was successful
        if result.returncode == 0:
            # Try to find the output file
            json_output_path = f"{output_name}.json"
            full_json_output_path = os.path.join(self.temp_dir, json_output_path)
            current_dir_json_path = os.path.join(os.getcwd(), json_output_path)

            # Check both possible locations for the output file
            extraction_result = None

            # Check all possible locations for the output file
            possible_paths = [
                # Temp directory
                full_json_output_path,
                # Current directory
                current_dir_json_path,
                # Project root directory
                os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{output_name}.json"),
                # Direct path without directory
                f"{output_name}.json",
                # Path with base filename only
                os.path.join(os.getcwd(), f"{base_filename}.json"),
                # Path with base filename and result suffix
                os.path.join(os.getcwd(), f"{base_filename}_result.json")
            ]

            # Debug all possible paths
            print("Checking for result file in the following locations:")
            for path in possible_paths:
                print(f"- {path} (exists: {os.path.exists(path)})")

            # Try to load from any of the possible locations
            for path in possible_paths:
                if os.path.exists(path):
                    try:
                        with open(path, 'r', encoding='utf-8') as f:
                            file_content = f.read().strip()
                            if file_content:
                                extraction_result = json.loads(file_content)
                                print(f"Successfully loaded result from {path}")
                                break
                    except json.JSONDecodeError as e:
                        print(f"Error parsing JSON from {path}: {str(e)}")
                        # Try to read the file content for debugging
                        try:
                            with open(path, 'r', encoding='utf-8') as f:
                                content = f.read()
                                print(f"File content: {content[:200]}...")
                        except Exception:
                            pass

            # If still not found, try to extract from stdout
            if extraction_result is None:
                # Try to extract result from stdout
                info_root_match = re.search(r'INFO:root:\s+(\{.*\})\s*$', result.stdout, re.MULTILINE)
                if info_root_match:
                    json_str = info_root_match.group(1)
                    try:
                        extraction_result = json.loads(json_str)
                        print(f"Successfully extracted result from stdout")
                    except json.JSONDecodeError as e:
                        print(f"Error parsing JSON from stdout: {str(e)}")

            # Handle list format
            if isinstance(extraction_result, list) and len(extraction_result) > 0:
                # If it's a list, use the first item
                extraction_result = extraction_result[0]
                print(f"Extracted result is a list, using first item")

            # Check if we found a result
            if extraction_result is not None:
                # Handle both dictionary and list formats
                if isinstance(extraction_result, dict):
                    # Check if the dictionary has any meaningful content
                    if extraction_result and any(k for k in extraction_result.keys() if k not in ['filename']):
                        print(f"Successfully extracted data: {json.dumps(extraction_result, default=str)[:200]}...")
                        return True, extraction_result, template_id, warnings
                    else:
                        print("Extraction result is an empty dictionary")
                        return False, None, None, f"Extraction result is an empty dictionary. Warnings: {warnings}"
                elif isinstance(extraction_result, list) and len(extraction_result) > 0:
                    # If it's a list, use the first item
                    first_item = extraction_result[0]
                    if isinstance(first_item, dict) and any(k for k in first_item.keys() if k not in ['filename']):
                        print(f"Successfully extracted data (from list): {json.dumps(first_item, default=str)[:200]}...")
                        return True, first_item, template_id, warnings
                    else:
                        print("First item in extraction result list is empty or invalid")
                        return False, None, None, f"First item in extraction result list is empty or invalid. Warnings: {warnings}"
                else:
                    print(f"Extraction result has unexpected type: {type(extraction_result)}")
                    return False, None, None, f"Extraction result has unexpected type: {type(extraction_result)}. Warnings: {warnings}"

            # Try to extract from stdout directly
            stdout_lines = result.stdout.splitlines()
            for line in stdout_lines:
                if "INFO:root:" in line and "{" in line:
                    try:
                        # Extract the JSON part
                        json_part = line.split("INFO:root:", 1)[1].strip()
                        data = json.loads(json_part)
                        if isinstance(data, dict) and any(k for k in data.keys() if k not in ['filename']):
                            print(f"Successfully extracted data from stdout: {json.dumps(data, default=str)[:200]}...")
                            return True, data, template_id, warnings
                    except (json.JSONDecodeError, IndexError):
                        pass

            # If we can't find the result, return failure
            print("Could not find extraction result")
            return False, None, None, f"Could not find extraction result. Warnings: {warnings}"
        else:
            # Command failed
            return False, None, None, f"Command failed with return code {result.returncode}. Stderr: {result.stderr}. Warnings: {warnings}"

    # Keep this method for backward compatibility, but make it use the selected template
    def process_pdf_with_all_templates(self, pdf_path):
        """Process a PDF with the selected template (for backward compatibility)

        This method now uses the selected template from the dropdown instead of trying all templates.

        Args:
            pdf_path (str): Path to the PDF file to process

        Returns:
            tuple: (success, result, template_id, warnings)
        """
        print("Note: Using selected template instead of trying all templates")
        return self.process_pdf_with_selected_template(pdf_path)

    def fetch_template_from_database(self, template_id):
        """Fetch a single template from the database by ID

        Args:
            template_id (int): ID of the template to fetch

        Returns:
            dict: Template data as a dictionary, or None if not found
        """
        try:
            print(f"[DEBUG] Fetching template with ID {template_id} from database")

            # Connect to database
            conn = sqlite3.connect("invoice_templates.db")
            cursor = conn.cursor()

            # First check if the template exists
            cursor.execute("SELECT COUNT(*) FROM templates WHERE id = ?", (template_id,))
            count = cursor.fetchone()[0]
            if count == 0:
                print(f"[DEBUG] Template with ID {template_id} not found in database")
                conn.close()
                return None

            # Fetch all template data to see what's available
            cursor.execute(
                """
                SELECT id, name, json_template
                FROM templates WHERE id = ?
                """,
                (template_id,),
            )
            template_row = cursor.fetchone()

            print(f"[DEBUG] Template row: ID={template_row[0]}, Name={template_row[1]}, Has JSON Template={bool(template_row[2])}")

            if not template_row[2]:
                print(f"[DEBUG] Template with ID {template_id} exists but has no JSON template")

                # Let's fetch the full template data to see what we have
                cursor.execute(
                    """
                    SELECT * FROM templates WHERE id = ?
                    """,
                    (template_id,),
                )
                full_row = cursor.fetchone()
                column_names = [description[0] for description in cursor.description]

                print(f"[DEBUG] Full template data columns: {column_names}")
                print(f"[DEBUG] Template has these fields: {[col for i, col in enumerate(column_names) if full_row[i] is not None]}")

                # Create a basic template structure using factory
                print(f"[DEBUG] Creating a basic template structure using factory since no JSON template exists")
                base_template = TemplateFactory.create_default_invoice_template(template_row[1], "INR")
                template_data = {
                    "issuer": base_template["issuer"],
                    "fields": {
                        "invoice_number": {
                            "parser": "regex",
                            "regex": "invoice_number:\\s*([^\\n]+)"
                        },
                        "date": {
                            "parser": "regex",
                            "regex": "date:\\s*([^\\n]+)"
                        },
                        "amount": {
                            "parser": "regex",
                            "regex": "amount:\\s*([^\\n]+)"
                        }
                    },
                    "keywords": base_template["keywords"],
                    "options": base_template["options"]
                }

                conn.close()
                return {"json_template": template_data}

            conn.close()

            # Parse JSON template
            try:
                template_data = json.loads(template_row[2])
                print(f"[DEBUG] Successfully parsed JSON template: {json.dumps(template_data, default=str)[:200]}...")
                return {"json_template": template_data}
            except json.JSONDecodeError as e:
                print(f"[DEBUG] Error parsing JSON template: {str(e)}")
                print(f"[DEBUG] Raw template data: {template_row[2][:200]}...")
                return None

        except Exception as e:
            print(f"[DEBUG] Error fetching template from database: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def test_invoice2data_template(self, pdf_path, template_id=None, template_data=None):
        """Test an invoice2data template against a PDF file

        Args:
            pdf_path (str): Path to the PDF file to test
            template_id (int, optional): ID of the template in the database. Required if template_data is None.
            template_data (dict, optional): The template data to use. If None, fetches from database using template_id.

        Returns:
            tuple: (success, result, warnings) where:
                - success (bool): True if extraction was successful
                - result (dict): The extracted data or None if extraction failed
                - warnings (str): Any warnings generated during extraction
        """
        # Import required modules
        import re
        import subprocess
        import io
        import logging

        if not os.path.exists(pdf_path):
            print(f"PDF file does not exist: {pdf_path}")
            return False, None, "PDF file does not exist"

        try:
            # Create a temporary directory for our files
            if not self.temp_dir:
                self.setup_temp_directory()

            # Get template data from database if not provided
            if template_data is None:
                if template_id is None:
                    return False, None, "Either template_id or template_data must be provided"

                # Fetch template from database
                template_data = self.fetch_template_from_database(template_id)

                if template_data is None:
                    return False, None, f"Template with ID {template_id} not found or has no JSON template"

            # Get a base filename from the PDF
            base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
            # Clean the filename to remove invalid characters
            base_filename = re.sub(r'[^\w\-_.]', '_', base_filename)

            # Extract just the JSON template part
            json_template = template_data.get("json_template", {})
            if not json_template:
                return False, None, "Template has no JSON template data"

            # Ensure the template has the required fields
            if "fields" not in json_template:
                json_template["fields"] = {}

            # Make sure the template has the three mandatory fields: amount, date, and invoice_number
            for field in ["amount", "date", "invoice_number"]:
                if field not in json_template["fields"]:
                    if field == "invoice_number":
                        json_template["fields"][field] = {
                            "parser": "regex",
                            "regex": "Invoice\\s+Number\\s*:?\\s*([\\w\\-\\/]+)"
                        }
                    elif field == "date":
                        json_template["fields"][field] = {
                            "parser": "regex",
                            "regex": "Date\\s*:?\\s*(\\d{1,2}[\\/\\-\\.]\\d{1,2}[\\/\\-\\.]\\d{2,4})",
                            "type": "date"
                        }
                    elif field == "amount":
                        json_template["fields"][field] = {
                            "parser": "regex",
                            "regex": "GRAND\\s+TOTAL\\s*[\\$\\€\\£]?\\s*(\\d+[\\.,]\\d+)",
                            "type": "float"
                        }

            # Add keywords if missing
            if "keywords" not in json_template or not json_template["keywords"]:
                # Use the base filename as a keyword
                json_template["keywords"] = [base_filename]

            # Add exclude_keywords if missing (required by InvoiceTemplate)
            if "exclude_keywords" not in json_template:
                json_template["exclude_keywords"] = []

            # Add date_formats if missing
            if "options" not in json_template:
                json_template["options"] = {
                    "currency": "INR",
                    "languages": ["en"],
                    "decimal_separator": ".",
                    "date_formats": [
                        '%d/%m/%Y',
                        '%d-%m-%Y',
                        '%Y-%m-%d',
                        '%d.%m.%Y'
                    ]
                }
            elif "date_formats" not in json_template.get("options", {}) or not json_template["options"].get("date_formats"):
                if "options" not in json_template:
                    json_template["options"] = {}
                json_template["options"]["date_formats"] = [
                    '%d/%m/%Y',
                    '%d-%m-%Y',
                    '%Y-%m-%d',
                    '%d.%m.%Y'
                ]

            # Create a common name for both text file and template
            common_name = f"temp_{base_filename}"

            # Create temporary template file
            temp_template_path = os.path.join(self.temp_dir, f"{common_name}.yml")
            print(f"Using template path: {temp_template_path}")

            # Add required fields for invoice2data
            if "name" not in json_template:
                json_template["name"] = common_name
            if "template_name" not in json_template:
                json_template["template_name"] = common_name

            # Save as YAML (preferred format)
            try:
                import yaml
                with open(temp_template_path, "w", encoding="utf-8") as f:
                    yaml.dump(json_template, f, default_flow_style=False, allow_unicode=True)
            except ImportError:
                # Fallback to JSON if YAML is not available
                temp_template_path = os.path.join(self.temp_dir, f"{common_name}.json")
                print(f"YAML module not available, using JSON instead: {temp_template_path}")
                with open(temp_template_path, "w", encoding="utf-8") as f:
                    json.dump(json_template, f, indent=2, ensure_ascii=False)

            # Template already saved above

            # Get the extracted data from the PDF
            extracted_data = self.processed_data.get(pdf_path, {})

            # Convert the extracted data to text
            pdf_text = self.convert_extraction_to_text(pdf_path, extracted_data)

            # Create temporary text file with the same common name
            temp_text_path = os.path.join(self.temp_dir, f"{common_name}.txt")
            with open(temp_text_path, "w", encoding="utf-8") as f:
                f.write(pdf_text)
            print(f"Saved extracted text to {temp_text_path}")

            # Define output name (without extension)
            output_name = f"{base_filename}_test_result"

            # Set up logging to capture warnings
            invoice2data_logger = logging.getLogger("invoice2data")
            log_capture = io.StringIO()
            string_handler = logging.StreamHandler(log_capture)
            string_handler.setLevel(logging.WARNING)
            invoice2data_logger.addHandler(string_handler)

            # Build the command
            cmd = [
                "invoice2data",
                "--input-reader", "text",
                "--exclude-built-in-templates",
                "--template-folder", self.temp_dir,
                "--output-format", "json",
                "--output-name", os.path.join(self.temp_dir, output_name),
                "--debug",
                temp_text_path
            ]

            # Try to use the InvoiceTemplate class directly first
            try:
                print("[DEBUG] Trying to use InvoiceTemplate directly...")
                from invoice2data.extract.invoice_template import InvoiceTemplate
                from invoice2data import extract_data

                # Create the template object directly
                template_obj = InvoiceTemplate(json_template)
                print(f"[DEBUG] Created InvoiceTemplate object: {template_obj.name}")

                # Print template details
                print(f"[DEBUG] Template details:")
                print(f"  - Name: {getattr(template_obj, 'name', 'N/A')}")
                print(f"  - Template name: {getattr(template_obj, 'template_name', 'N/A')}")
                print(f"  - Keywords: {getattr(template_obj, 'keywords', 'N/A')}")
                print(f"  - Fields: {list(getattr(template_obj, 'fields', {}).keys())}")

                # Extract data using the API directly
                api_result = extract_data(temp_text_path, templates=[template_obj])
                if api_result:
                    print(f"[DEBUG] Successfully extracted data using API: {json.dumps(api_result, default=str)[:200]}...")
                    return True, api_result, ""
                else:
                    print("[DEBUG] API extraction returned no results, falling back to command-line approach")
            except Exception as api_error:
                print(f"[DEBUG] Error using API directly: {str(api_error)}")
                print("[DEBUG] Falling back to command-line approach")

            # Run the command and capture output as fallback
            print("[DEBUG] Using command-line approach...")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                check=False
            )

            # Get warnings
            warnings = log_capture.getvalue().strip()

            # Check if command was successful
            if result.returncode == 0:
                # Try to find the output file
                json_output_path = f"{output_name}.json"
                full_json_output_path = os.path.join(self.temp_dir, json_output_path)
                current_dir_json_path = os.path.join(os.getcwd(), json_output_path)

                # Check both possible locations for the output file
                extraction_result = None

                # Check all possible locations for the output file
                possible_paths = [
                    # Temp directory
                    full_json_output_path,
                    # Current directory
                    current_dir_json_path,
                    # Project root directory
                    os.path.join(os.path.dirname(os.path.abspath(__file__)), f"{output_name}.json"),
                    # Direct path without directory
                    f"{output_name}.json",
                    # Path with base filename only
                    os.path.join(os.getcwd(), f"{base_filename}.json"),
                    # Path with base filename and result suffix
                    os.path.join(os.getcwd(), f"{base_filename}_result.json")
                ]

                # Debug all possible paths
                print("Checking for result file in the following locations:")
                for path in possible_paths:
                    print(f"- {path} (exists: {os.path.exists(path)})")

                # Try to load from any of the possible locations
                for path in possible_paths:
                    if os.path.exists(path):
                        try:
                            with open(path, 'r', encoding='utf-8') as f:
                                file_content = f.read().strip()
                                if file_content:
                                    extraction_result = json.loads(file_content)
                                    print(f"Successfully loaded result from {path}")
                                    break
                        except json.JSONDecodeError as e:
                            print(f"Error parsing JSON from {path}: {str(e)}")
                            # Try to read the file content for debugging
                            try:
                                with open(path, 'r', encoding='utf-8') as f:
                                    content = f.read()
                                    print(f"File content: {content[:200]}...")
                            except Exception:
                                pass

                # If still not found, try to extract from stdout
                if extraction_result is None:
                    # Try to extract result from stdout
                    info_root_match = re.search(r'INFO:root:\s+(\{.*\})\s*$', result.stdout, re.MULTILINE)
                    if info_root_match:
                        json_str = info_root_match.group(1)
                        try:
                            extraction_result = json.loads(json_str)
                            print(f"Successfully extracted result from stdout")
                        except json.JSONDecodeError as e:
                            print(f"Error parsing JSON from stdout: {str(e)}")

                # Check if we found a result
                if extraction_result is not None:
                    # Handle both dictionary and list formats
                    if isinstance(extraction_result, dict):
                        # Check if the dictionary has any meaningful content
                        if extraction_result and any(k for k in extraction_result.keys() if k not in ['filename']):
                            print(f"Successfully extracted data: {json.dumps(extraction_result, default=str)[:200]}...")
                            return True, extraction_result, warnings
                        else:
                            print("Extraction result is an empty dictionary")
                            return False, None, f"Extraction result is an empty dictionary. Warnings: {warnings}"
                    elif isinstance(extraction_result, list) and len(extraction_result) > 0:
                        # If it's a list, use the first item
                        first_item = extraction_result[0]
                        if isinstance(first_item, dict) and any(k for k in first_item.keys() if k not in ['filename']):
                            print(f"Successfully extracted data (from list): {json.dumps(first_item, default=str)[:200]}...")
                            return True, first_item, warnings
                        else:
                            print("First item in extraction result list is empty or invalid")
                            return False, None, f"First item in extraction result list is empty or invalid. Warnings: {warnings}"
                    elif isinstance(extraction_result, list) and len(extraction_result) == 0:
                        # Empty list means no results
                        print("Extraction returned an empty list")
                        return False, None, f"Extraction returned an empty list. Warnings: {warnings}"
                    else:
                        print(f"Extraction result has unexpected type: {type(extraction_result)}")
                        return False, None, f"Extraction result has unexpected type: {type(extraction_result)}. Warnings: {warnings}"

                # Try to extract from stdout directly
                stdout_lines = result.stdout.splitlines()
                for line in stdout_lines:
                    if "INFO:root:" in line and "{" in line:
                        try:
                            # Extract the JSON part
                            json_part = line.split("INFO:root:", 1)[1].strip()
                            data = json.loads(json_part)
                            if isinstance(data, dict) and any(k for k in data.keys() if k not in ['filename']):
                                print(f"Successfully extracted data from stdout: {json.dumps(data, default=str)[:200]}...")
                                return True, data, warnings
                        except (json.JSONDecodeError, IndexError):
                            pass

                # If we can't find the result, return failure
                print("Could not find extraction result")
                return False, None, f"Could not find extraction result. Warnings: {warnings}"
            else:
                # Command failed
                return False, None, f"Command failed with return code {result.returncode}. Stderr: {result.stderr}. Warnings: {warnings}"

        except Exception as e:
            print(f"Error testing invoice2data template: {str(e)}")
            import traceback
            traceback.print_exc()
            return False, None, f"Error: {str(e)}"

    def process_pdfs_with_template(self, pdf_paths, template_id):
        """Process multiple PDFs with a specific template

        Args:
            pdf_paths (list): List of paths to PDF files to process
            template_id (int): ID of the template to use

        Returns:
            dict: Dictionary mapping PDF paths to extraction results
        """
        results = {}

        # Fetch template from database
        template_data = self.fetch_template_from_database(template_id)

        if template_data is None:
            print(f"Template with ID {template_id} not found or has no JSON template")
            return results

        # Process each PDF
        for pdf_path in pdf_paths:
            print(f"Processing {pdf_path} with template ID {template_id}")

            # Test the template against the PDF
            success, result, warnings = self.test_invoice2data_template(
                pdf_path,
                template_data=template_data
            )

            if success:
                results[pdf_path] = result
                print(f"Successfully extracted data from {pdf_path}")
            else:
                results[pdf_path] = None
                print(f"Failed to extract data from {pdf_path}: {warnings}")

        return results



    def process_selected_pdfs_with_template(self, template_id=None):
        """Process selected PDFs with a specific template

        This method is intended to be called from a button click or menu action.

        Args:
            template_id (int, optional): ID of the template to use. If None, prompts the user to select a template.

        Returns:
            dict: Dictionary mapping PDF paths to extraction results
        """
        # Get selected PDF paths
        selected_pdfs = []
        for index in range(self.file_list.count()):
            item = self.file_list.item(index)
            if item.checkState() == Qt.Checked:
                pdf_path = item.data(Qt.UserRole)
                selected_pdfs.append(pdf_path)

        if not selected_pdfs:
            QMessageBox.warning(self, "No PDFs Selected", "Please select at least one PDF to process.")
            return {}

        # If no template ID provided, prompt user to select one
        if template_id is None:
            # Connect to database
            conn = sqlite3.connect("invoice_templates.db")
            cursor = conn.cursor()

            # Fetch all templates
            cursor.execute("SELECT id, name FROM templates")
            templates = cursor.fetchall()
            conn.close()

            if not templates:
                QMessageBox.warning(self, "No Templates", "No templates found in the database.")
                return {}

            # Create a dialog to select a template
            dialog = QDialog(self)
            dialog.setWindowTitle("Select Template")
            layout = QVBoxLayout()

            # Add a combo box with templates
            combo = QComboBox()
            for template_id, template_name in templates:
                combo.addItem(template_name, template_id)
            layout.addWidget(combo)

            # Add buttons
            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addWidget(buttons)

            dialog.setLayout(layout)

            # Show dialog
            if dialog.exec_() == QDialog.Accepted:
                template_id = combo.currentData()
            else:
                return {}

        # Process PDFs with template
        results = self.process_pdfs_with_template(selected_pdfs, template_id)

        # Show results
        success_count = sum(1 for result in results.values() if result is not None)
        QMessageBox.information(
            self,
            "Processing Complete",
            f"Processed {len(results)} PDFs with template ID {template_id}.\n"
            f"Successfully extracted data from {success_count} PDFs."
        )

        # Update the UI to show extraction results
        self.update_extraction_results(results)

        return results

    def update_extraction_results(self, results):
        """Update the UI to show extraction results

        Args:
            results (dict): Dictionary mapping PDF paths to extraction results
        """
        # Update the file list to show extraction status
        for index in range(self.file_list.count()):
            item = self.file_list.item(index)
            pdf_path = item.data(Qt.UserRole)

            if pdf_path in results:
                result = results[pdf_path]
                if result is not None:
                    # Extraction successful
                    item.setBackground(QColor(200, 255, 200))  # Light green
                    item.setToolTip(f"Extraction successful: {json.dumps(result, indent=2)}")
                else:
                    # Extraction failed
                    item.setBackground(QColor(255, 200, 200))  # Light red
                    item.setToolTip("Extraction failed")

        # If we have a currently selected file, update the display
        if self.current_file:
            if self.current_file in results and results[self.current_file] is not None:
                # Show the extracted data
                self.display_extraction_result(results[self.current_file])

    def display_extraction_result(self, result):
        """Display the extraction result in the UI

        Args:
            result (dict): The extraction result from invoice2data
        """
        # Clear existing data
        self.clear_data_display()

        # Create a formatted string with the extraction result
        result_text = json.dumps(result, indent=2)

        # Display the result in a text edit
        result_display = QTextEdit()
        result_display.setReadOnly(True)
        result_display.setFont(QFont("Courier New", 10))
        result_display.setText(result_text)

        # Add to layout
        result_label = QLabel("Extraction Result:")
        self.right_layout.addWidget(result_label)
        self.right_layout.addWidget(result_display)

    def _extract_with_dual_coordinates(self, pdf_path, template_data):
        """Extract using dual coordinate system - no coordinate conversion needed"""
        try:
            print("\n" + "=" * 80)
            print(f"DUAL COORDINATE EXTRACTION")
            print("=" * 80)

            # Load the PDF document
            print(f"Loading PDF document: {pdf_path}")
            pdf_document = fitz.open(pdf_path)
            pdf_page_count = len(pdf_document)
            print(f"✓ PDF document loaded successfully with {pdf_page_count} pages")

            # Initialize results dictionary
            results = {
                "header_tables": [],
                "items_tables": [],
                "summary_tables": [],
                "extraction_status": {
                    "header": "not_processed",
                    "items": "not_processed",
                    "summary": "not_processed",
                    "overall": "not_processed"
                }
            }

            # Get extraction regions based on template type
            if template_data["template_type"] == "single":
                # Single-page template - use extraction_regions directly
                extraction_regions = template_data.get("extraction_regions", {})
                extraction_column_lines = template_data.get("extraction_column_lines", {})

                print(f"Single-page template: processing page 1")
                self._process_page_with_dual_coords(
                    pdf_document, 0, extraction_regions, extraction_column_lines,
                    template_data, results, pdf_path
                )

            else:
                # Multi-page template - use extraction_page_regions
                extraction_page_regions = template_data.get("extraction_page_regions", [])
                extraction_page_column_lines = template_data.get("extraction_page_column_lines", [])

                print(f"Multi-page template: processing {len(extraction_page_regions)} template pages")

                # Process each page
                for page_index in range(pdf_page_count):
                    # Map PDF page to template page (simplified mapping for now)
                    template_page_idx = min(page_index, len(extraction_page_regions) - 1)

                    if template_page_idx < len(extraction_page_regions):
                        page_regions = extraction_page_regions[template_page_idx]
                        page_column_lines = extraction_page_column_lines[template_page_idx] if template_page_idx < len(extraction_page_column_lines) else {}

                        print(f"Processing PDF page {page_index + 1} using template page {template_page_idx + 1}")
                        self._process_page_with_dual_coords(
                            pdf_document, page_index, page_regions, page_column_lines,
                            template_data, results, pdf_path
                        )

            # Close the PDF document
            pdf_document.close()

            # Update overall extraction status
            if results["extraction_status"]["items"] == "success":
                results["extraction_status"]["overall"] = "success"
            elif any(status == "success" for status in [results["extraction_status"]["header"], results["extraction_status"]["summary"]]):
                results["extraction_status"]["overall"] = "partial"
            else:
                results["extraction_status"]["overall"] = "failed"

            print(f"\nDual coordinate extraction summary:")
            print(f"  Header: {results['extraction_status']['header']}")
            print(f"  Items: {results['extraction_status']['items']}")
            print(f"  Summary: {results['extraction_status']['summary']}")
            print(f"  Overall: {results['extraction_status']['overall']}")

            return results

        except Exception as e:
            handle_exception(
                func_name="_extract_with_dual_coordinates",
                exception=e,
                context={"pdf_path": pdf_path}
            )
            if "pdf_document" in locals():
                pdf_document.close()
            return None

    def _process_page_with_dual_coords(self, pdf_document, page_index, extraction_regions, extraction_column_lines, template_data, results, pdf_path):
        """Process a single page using dual coordinate regions"""
        try:
            from dual_coordinate_storage import DualCoordinateRegion, DualCoordinateColumnLine

            print(f"\nProcessing page {page_index + 1} with dual coordinates")

            # Get config parameters
            config = template_data.get("config", {})

            # Process each section (header, items, summary)
            for section in ["header", "items", "summary"]:
                if section in extraction_regions and extraction_regions[section]:
                    dual_regions = extraction_regions[section]
                    dual_column_lines = extraction_column_lines.get(section, [])

                    print(f"\nExtracting {section} section from page {page_index + 1}")
                    print(f"  Found {len(dual_regions)} dual coordinate region(s)")
                    print(f"  Found {len(dual_column_lines)} dual coordinate column line(s)")

                    # Extract coordinates directly from dual coordinate regions
                    table_areas = []
                    columns_list = []

                    for region_idx, dual_region in enumerate(dual_regions):
                        # Get extraction coordinates directly - no conversion needed!
                        extraction_coords = dual_region.get_extraction_coordinates()
                        table_areas.append(extraction_coords)

                        print(f"  Region {region_idx}: {dual_region.label} -> {extraction_coords}")

                        # Get column lines for this region
                        region_columns = []
                        for dual_line in dual_column_lines:
                            # Get extraction coordinates for column lines
                            line_coords = dual_line.get_extraction_coordinates()
                            # Use x-coordinate of the line
                            region_columns.append(line_coords[0])  # start_x

                        # Format column lines
                        col_str = ",".join([str(x) for x in sorted(region_columns)]) if region_columns else ""
                        columns_list.append(col_str)

                    # Get extraction parameters
                    extraction_params = config.get("extraction_params", {})
                    section_params = extraction_params.get(section, {})
                    row_tol = section_params.get("row_tol", 5)  # Default row tolerance

                    # Prepare extraction parameters
                    normalized_params = {
                        section: {"row_tol": row_tol},
                        "split_text": extraction_params.get("split_text", True),
                        "strip_text": extraction_params.get("strip_text", "\n"),
                        "flavor": extraction_params.get("flavor", "stream"),
                        "parallel": True
                    }

                    print(f"  Extraction parameters: row_tol={row_tol}")
                    print(f"  Table areas: {table_areas}")
                    print(f"  Columns: {columns_list}")

                    # Extract tables using the direct coordinates with multi-method support
                    try:
                        extraction_method = template_data.get("extraction_method", "pypdf_table_extraction")

                        if len(table_areas) > 1:
                            # Multiple tables - use multi-method extraction
                            table_dfs = extract_with_method(
                                pdf_path=pdf_path,
                                extraction_method=extraction_method,
                                page_number=page_index + 1,
                                table_areas=table_areas,
                                columns_list=columns_list,
                                section_type=section,
                                extraction_params=normalized_params,
                                use_cache=True
                            )

                            # Process each table
                            for table_index, table_df in enumerate(table_dfs or []):
                                if table_df is not None and not table_df.empty:
                                    self._process_extracted_table(
                                        table_df, section, results,
                                        region_index=table_index,
                                        page_number=page_index + 1
                                    )
                        else:
                            # Single table - use multi-method extraction
                            table_df = extract_with_method(
                                pdf_path=pdf_path,
                                extraction_method=extraction_method,
                                page_number=page_index + 1,
                                table_areas=table_areas,
                                columns_list=columns_list,
                                section_type=section,
                                extraction_params=normalized_params,
                                use_cache=True
                            )

                            if table_df is not None and not table_df.empty:
                                self._process_extracted_table(
                                    table_df, section, results,
                                    region_index=0,
                                    page_number=page_index + 1
                                )

                        # Update extraction status
                        results["extraction_status"][section] = "success"

                    except Exception as e:
                        print(f"  Error extracting {section}: {e}")
                        results["extraction_status"][section] = "failed"

                else:
                    print(f"  No {section} regions defined for page {page_index + 1}")

        except Exception as e:
            handle_exception(
                func_name="_process_page_with_dual_coords",
                exception=e,
                context={"page": page_index + 1, "pdf_path": pdf_path}
            )

