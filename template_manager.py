from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QScrollArea, QFrame, QLineEdit, QTextEdit,
                             QTableWidget, QTableWidgetItem, QHeaderView, QDialog,
                             QFormLayout, QMessageBox, QInputDialog, QDialogButtonBox,
                             QApplication, QTabWidget, QCheckBox, QComboBox, QGroupBox, QGridLayout, QSpinBox, QListWidget, QListWidgetItem,
                             QMainWindow, QStackedWidget, QFileDialog, QScrollArea, QFrame, QSplitter, QGridLayout, QLineEdit, QComboBox,
                             QListWidget, QProgressBar, QTabWidget, QTextEdit, QCheckBox, QProgressDialog, QMenu)
from PySide6.QtCore import Qt, Signal, QRect, QPoint
from PySide6.QtGui import QFont, QIcon, QColor
from database import InvoiceDatabase
import os
import json
import sqlite3
from datetime import datetime

# Import new factory modules for code deduplication
from common_factories import (
    TemplateFactory, DatabaseOperationFactory, UIMessageFactory,
    ValidationFactory, get_database_factory
)
from ui_component_factory import UIComponentFactory, LayoutFactory

class SaveTemplateDialog(QDialog):
    """Dialog for saving a new template"""

    def __init__(self, parent=None, template_name=None, template_description=None):
        super().__init__(parent)
        self.setWindowTitle("Save Invoice Template")
        self.setMinimumWidth(450)  # Make the dialog wider

        # Set global style for this dialog
        self.setStyleSheet("""
            QDialog {
                background-color: white;
            }
            QLabel {
                color: black;
            }
            QLineEdit, QTextEdit {
                color: black;
                background-color: white;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
            }
            QFormLayout QLabel {
                color: black;
                font-weight: bold;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # Add header/title
        header_label = QLabel("Template Information")
        header_label.setFont(QFont("Arial", 14, QFont.Bold))
        header_label.setStyleSheet("color: #333;")
        layout.addWidget(header_label)

        # Add description
        description_label = QLabel("Fill in the details to save your invoice template for future use.")
        description_label.setWordWrap(True)
        description_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(description_label)

        form_layout = QFormLayout()
        form_layout.setSpacing(12)

        # Explicitly set label styling for form fields
        name_label = QLabel("Template Name:")
        name_label.setStyleSheet("color: black; font-weight: bold;")

        desc_label = QLabel("Description:")
        desc_label.setStyleSheet("color: black; font-weight: bold;")

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Enter a descriptive name for your template")
        self.name_input.setStyleSheet("color: black; background-color: white;")
        if template_name:
            self.name_input.setText(template_name)

        self.description_input = QTextEdit()
        self.description_input.setPlaceholderText("Describe the purpose or usage of this template (optional)")
        self.description_input.setMinimumHeight(100)
        self.description_input.setStyleSheet("color: black; background-color: white;")
        if template_description:
            self.description_input.setText(template_description)

        form_layout.addRow(name_label, self.name_input)
        form_layout.addRow(desc_label, self.description_input)

        layout.addLayout(form_layout)

        # Buttons
        buttons_layout = QHBoxLayout()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                color: #212529;
                padding: 10px 20px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-weight: normal;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                color: black;
            }
        """)

        save_btn = QPushButton("Save Template")
        save_btn.clicked.connect(self.accept)
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 10px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3158D3;
                color: white;
            }
        """)
        save_btn.setDefault(True)

        buttons_layout.addStretch(1)
        buttons_layout.addWidget(cancel_btn)
        buttons_layout.addWidget(save_btn)

        layout.addLayout(buttons_layout)

        self.setLayout(layout)

    def get_template_data(self):
        """Get the template name and description entered by the user"""
        return {
            "name": self.name_input.text().strip(),
            "description": self.description_input.toPlainText().strip()
        }

class EditTemplateDialog(QDialog):
    """Dialog for editing template settings and configuration"""

    def __init__(self, parent=None, template_data=None):
        super().__init__(parent)
        self.template_data = template_data
        self.setWindowTitle("Edit Template Settings")
        self.setMinimumWidth(800)
        self.setMinimumHeight(600)

        # Initialize page count
        self.page_count = template_data.get('page_count', 1) if template_data else 1
        self.current_page = 0

        # Initialize validation rules
        self.validation_rules = template_data.get('validation_rules', {}) if template_data else {}

        # Initialize mapping configuration
        self.mapping_config = template_data.get('mapping_config', {
            'approach': 'page_wise',  # Default to page-wise mapping
            'page_wise': {
                'first_page': 1,
                'middle_pages': 'sequential',
                'last_page': 'last_template_page'
            },
            'region_wise': {
                'header': {'source_page': '1'},
                'items': {'source_page': '1-n'},
                'summary': {'source_page': 'n'}
            }
        }) if template_data else {
            'approach': 'page_wise',
            'page_wise': {
                'first_page': 1,
                'middle_pages': 'sequential',
                'last_page': 'last_template_page'
            },
            'region_wise': {
                'header': {'source_page': '1'},
                'items': {'source_page': '1-n'},
                'summary': {'source_page': 'n'}
            }
        }

        # Initialize regions attribute to avoid 'can't set attribute' error
        self.regions = {
            'header': [],
            'items': [],
            'summary': []
        }

        # Set global style for this dialog
        self.setStyleSheet("""
            QDialog {
                background-color: white;
            }
            QLabel {
                color: black;
            }
            QLineEdit, QTextEdit {
                color: black;
                background-color: white;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
            }
            QFormLayout QLabel {
                color: black;
                font-weight: bold;
            }
        """)

        # Create main layout
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Add header
        header_label = QLabel("Template Settings")
        header_label.setFont(QFont("Arial", 16, QFont.Bold))
        header_label.setStyleSheet("color: #333;")
        main_layout.addWidget(header_label)

        # Create tab widget for different sections
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
            }
            QTabBar::tab {
                background: #f0f0f0;
                border: 1px solid #ddd;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                padding: 8px 12px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: none;
                margin-bottom: -1px;
            }
        """)

        # Add tabs
        tab_widget.addTab(self.create_general_tab(), "General")
        tab_widget.addTab(self.create_regions_tab(), "Table Regions")
        tab_widget.addTab(self.create_columns_tab(), "Column Lines")
        tab_widget.addTab(self.create_config_tab(), "Configuration")
        tab_widget.addTab(self.create_mapping_tab(), "Mapping")
        tab_widget.addTab(self.create_validation_tab(), "YAML Template")

        main_layout.addWidget(tab_widget)

        # Add buttons
        buttons_layout = QHBoxLayout()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f8f9fa;
                color: #212529;
                padding: 10px 20px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-weight: normal;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                color: black;
            }
        """)

        save_btn = QPushButton("Save Changes")
        save_btn.clicked.connect(self.accept)
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 10px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3158D3;
                color: white;
            }
        """)
        save_btn.setDefault(True)

        buttons_layout.addStretch(1)
        buttons_layout.addWidget(cancel_btn)
        buttons_layout.addWidget(save_btn)

        main_layout.addLayout(buttons_layout)

        self.setLayout(main_layout)

    def create_general_tab(self):
        """Create the general settings tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(16)

        # Template name
        name_layout = QFormLayout()
        self.name_input = QLineEdit()
        self.name_input.setText(self.template_data.get("name", ""))
        name_layout.addRow("Template Name:", self.name_input)

        # Description
        desc_layout = QFormLayout()
        self.desc_input = QTextEdit()
        self.desc_input.setMinimumHeight(100)
        self.desc_input.setText(self.template_data.get("description", ""))
        desc_layout.addRow("Description:", self.desc_input)

        # Template type
        type_layout = QFormLayout()
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Single-page", "Multi-page"])
        self.type_combo.setCurrentText("Multi-page" if self.template_data.get("template_type") == "multi" else "Single-page")
        self.type_combo.currentTextChanged.connect(self.on_template_type_changed)
        type_layout.addRow("Template Type:", self.type_combo)

        # Page count (only for multi-page)
        # Create a container widget for the page count layout
        self.page_count_container = QWidget()
        self.page_count_layout = QFormLayout(self.page_count_container)
        self.page_count_spin = QSpinBox()
        self.page_count_spin.setMinimum(1)
        self.page_count_spin.setMaximum(10)  # Reasonable limit
        self.page_count_spin.setValue(self.page_count)
        self.page_count_spin.valueChanged.connect(self.on_page_count_changed)
        self.page_count_layout.addRow("Number of Pages:", self.page_count_spin)

        # Show/hide page count based on template type
        self.update_page_count_visibility()

        layout.addLayout(name_layout)
        layout.addLayout(desc_layout)
        layout.addLayout(type_layout)
        layout.addWidget(self.page_count_container)
        layout.addStretch()

        return tab

    def create_regions_tab(self):
        """Create the table regions tab with page support"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(16)

        # Page navigation for multi-page templates
        if self.template_data.get("template_type") == "multi":
            page_nav_layout = QHBoxLayout()
            self.prev_page_btn = QPushButton("← Previous Page")
            self.next_page_btn = QPushButton("Next Page →")
            self.page_label = QLabel(f"Page {self.current_page + 1} of {self.page_count}")

            self.prev_page_btn.clicked.connect(self.prev_page)
            self.next_page_btn.clicked.connect(self.next_page)

            page_nav_layout.addWidget(self.prev_page_btn)
            page_nav_layout.addWidget(self.page_label)
            page_nav_layout.addWidget(self.next_page_btn)

            # Clone regions button removed

            layout.addLayout(page_nav_layout)

        # Create table for regions
        self.regions_table = QTableWidget()
        self.regions_table.setColumnCount(7)
        self.regions_table.setHorizontalHeaderLabels([
            "Section", "Table #", "X", "Y", "Width", "Height", "Scaled Format"
        ])
        self.regions_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.regions_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.regions_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.regions_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.regions_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.regions_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.regions_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)

        # Add regions from template data
        self.load_regions_for_current_page()

        layout.addWidget(self.regions_table)

        # Add buttons for region management
        buttons_layout = QHBoxLayout()
        add_btn = QPushButton("Add Region")
        edit_btn = QPushButton("Edit Region")
        delete_btn = QPushButton("Delete Region")

        buttons_layout.addWidget(add_btn)
        buttons_layout.addWidget(edit_btn)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addStretch()

        layout.addLayout(buttons_layout)

        return tab

    def create_columns_tab(self):
        """Create the column lines tab with page support"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(16)

        # Page navigation for multi-page templates
        if self.template_data.get("template_type") == "multi":
            page_nav_layout = QHBoxLayout()
            self.prev_page_btn_cols = QPushButton("← Previous Page")
            self.next_page_btn_cols = QPushButton("Next Page →")
            self.page_label_cols = QLabel(f"Page {self.current_page + 1} of {self.page_count}")

            self.prev_page_btn_cols.clicked.connect(self.prev_page)
            self.next_page_btn_cols.clicked.connect(self.next_page)

            page_nav_layout.addWidget(self.prev_page_btn_cols)
            page_nav_layout.addWidget(self.page_label_cols)
            page_nav_layout.addWidget(self.next_page_btn_cols)

            # Clone column lines button removed

            layout.addLayout(page_nav_layout)

        # Create table for column lines
        self.columns_table = QTableWidget()
        self.columns_table.setColumnCount(4)
        self.columns_table.setHorizontalHeaderLabels(["Section", "Table #", "X Position", "Description"])
        self.columns_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.columns_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.columns_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.columns_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)

        # Add column lines from template data
        self.load_column_lines_for_current_page()

        layout.addWidget(self.columns_table)

        # Add buttons for column line management
        buttons_layout = QHBoxLayout()
        add_btn = QPushButton("Add Column Line")
        edit_btn = QPushButton("Edit Column Line")
        delete_btn = QPushButton("Delete Column Line")

        buttons_layout.addWidget(add_btn)
        buttons_layout.addWidget(edit_btn)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addStretch()

        layout.addLayout(buttons_layout)

        return tab

    def create_config_tab(self):
        """Create the configuration tab with extraction parameters"""
        tab = QWidget()
        main_layout = QVBoxLayout(tab)
        main_layout.setSpacing(16)

        # Page navigation for multi-page templates with page-specific config
        if self.template_data.get("template_type") == "multi":
            page_nav_layout = QHBoxLayout()
            self.prev_page_btn_config = QPushButton("← Previous Page")
            self.next_page_btn_config = QPushButton("Next Page →")
            self.page_label_config = QLabel(f"Page {self.current_page + 1} of {self.page_count}")

            self.prev_page_btn_config.clicked.connect(self.prev_page)
            self.next_page_btn_config.clicked.connect(self.next_page)

            page_nav_layout.addWidget(self.prev_page_btn_config)
            page_nav_layout.addWidget(self.page_label_config)
            page_nav_layout.addWidget(self.next_page_btn_config)

            # Add page-specific config checkbox
            self.page_specific_config = QCheckBox("Enable Page-Specific Configuration")
            self.page_specific_config.setToolTip("Configure extraction parameters separately for each page")

            # Check if page-specific configs already exist
            if 'page_configs' in self.template_data:
                self.page_specific_config.setChecked(True)

            page_nav_layout.addWidget(self.page_specific_config)
            main_layout.addLayout(page_nav_layout)

        # Create sections using group boxes
        # 1. Extraction Parameters Section
        extraction_group = QGroupBox("Extraction Parameters")
        extraction_layout = QVBoxLayout()

        # Row tolerance parameters
        tol_form = QFormLayout()

        # Get default values from template data if available
        config = self.template_data.get('config', {})

        # Handle both old and new format - check for extraction_params first
        extraction_params = {}
        if 'extraction_params' in config:
            extraction_params = config.get('extraction_params', {})
            print("Found extraction_params in config - using nested structure")
        else:
            # Fallback to old format where parameters might be directly in config
            print("No extraction_params found - using direct config values")

        # Header row tolerance
        header_params = {}
        if 'header' in extraction_params:
            header_params = extraction_params.get('header', {})
        elif 'header' in config:
            header_params = config.get('header', {})

        self.header_row_tol = QSpinBox()
        self.header_row_tol.setRange(0, 999)  # Increased maximum value to 999 (effectively removing the constraint)
        row_tol_value = header_params.get('row_tol', None)
        if row_tol_value is None:
            row_tol_value = config.get('row_tol', 3)  # Global fallback
        self.header_row_tol.setValue(row_tol_value)
        self.header_row_tol.setToolTip("Tolerance for header row extraction (higher = more flexible)")
        tol_form.addRow("Header Row Tolerance:", self.header_row_tol)

        # Items row tolerance
        items_params = {}
        if 'items' in extraction_params:
            items_params = extraction_params.get('items', {})
        elif 'items' in config:
            items_params = config.get('items', {})

        self.items_row_tol = QSpinBox()
        self.items_row_tol.setRange(0, 999)  # Increased maximum value to 999 (effectively removing the constraint)
        row_tol_value = items_params.get('row_tol', None)
        if row_tol_value is None:
            row_tol_value = config.get('row_tol', 3)  # Global fallback
        self.items_row_tol.setValue(row_tol_value)
        self.items_row_tol.setToolTip("Tolerance for items row extraction (higher = more flexible)")
        tol_form.addRow("Items Row Tolerance:", self.items_row_tol)

        # Summary row tolerance
        summary_params = {}
        if 'summary' in extraction_params:
            summary_params = extraction_params.get('summary', {})
        elif 'summary' in config:
            summary_params = config.get('summary', {})

        self.summary_row_tol = QSpinBox()
        self.summary_row_tol.setRange(0, 999)  # Increased maximum value to 999 (effectively removing the constraint)
        row_tol_value = summary_params.get('row_tol', None)
        if row_tol_value is None:
            row_tol_value = config.get('row_tol', 3)  # Global fallback
        self.summary_row_tol.setValue(row_tol_value)
        self.summary_row_tol.setToolTip("Tolerance for summary row extraction (higher = more flexible)")
        tol_form.addRow("Summary Row Tolerance:", self.summary_row_tol)

        extraction_layout.addLayout(tol_form)

        # Text processing options
        text_form = QFormLayout()

        # Split text option
        self.split_text = QCheckBox("Enable")
        split_text_value = extraction_params.get('split_text', None)
        if split_text_value is None:
            split_text_value = config.get('split_text', True)
        self.split_text.setChecked(split_text_value)
        self.split_text.setToolTip("Split text that may contain multiple values")
        text_form.addRow("Split Text:", self.split_text)

        # Strip text option
        self.strip_text = QLineEdit()
        strip_text_value = extraction_params.get('strip_text', None)
        if strip_text_value is None:
            strip_text_value = config.get('strip_text', '\n')
        if strip_text_value == '\n':
            self.strip_text.setText("\\n")
        else:
            self.strip_text.setText(strip_text_value)
        self.strip_text.setToolTip("Characters to strip from text (use \\n for newlines)")
        text_form.addRow("Strip Text:", self.strip_text)

        extraction_layout.addLayout(text_form)

        # Multi-table mode option
        self.multi_table_mode = QCheckBox("Enable Multi-Table Mode")
        self.multi_table_mode.setChecked(config.get('multi_table_mode', False))
        self.multi_table_mode.setToolTip("Process multiple tables in the header section")
        extraction_layout.addWidget(self.multi_table_mode)

        extraction_group.setLayout(extraction_layout)
        main_layout.addWidget(extraction_group)

        # Multi-page options removed - using simplified page-wise approach
        # Page mapping features will be implemented later as a separate enhancement

        # Add additional parameters section if they exist
        if hasattr(self, 'template_data'):
            additional_params = []

            # Get the appropriate config based on template type and current page
            if self.template_data.get('template_type') == 'multi' and 'page_configs' in self.template_data:
                # For multi-page templates, show page-specific config if available
                page_configs = self.template_data['page_configs']
                if self.current_page < len(page_configs) and page_configs[self.current_page]:
                    config = page_configs[self.current_page]
                    print(f"[DEBUG] Using page-specific config for page {self.current_page + 1}: {list(config.keys())}")

                    # Also include global config parameters
                    if 'config' in self.template_data:
                        global_config = self.template_data['config']
                        print(f"[DEBUG] Also including global config parameters: {list(global_config.keys())}")

                        # Add a special parameter to indicate this is page-specific config
                        additional_params.append(("_page_specific_config", f"Page {self.current_page + 1} Configuration"))

                        # Add global parameters that aren't in page-specific config
                        for key, value in global_config.items():
                            if key not in config:
                                additional_params.append((f"global_{key}", value))
                else:
                    # If no page-specific config, use global config
                    config = self.template_data.get('config', {})
                    print(f"[DEBUG] No page-specific config for page {self.current_page + 1}, using global config")
            else:
                # For single-page templates, use the regular config
                config = self.template_data.get('config', {})
                print(f"[DEBUG] Using single-page template config: {list(config.keys())}")

            # Check for any custom parameters (exclude known parameters)
            known_params = ['multi_table_mode', 'extraction_params', 'regex_patterns', 'use_middle_page',
                           'fixed_page_count', 'total_pages', 'page_indices', 'store_original_coords',
                           'original_regions', 'original_column_lines', 'scale_factors']

            # Add parameters from the selected config
            for key, value in config.items():
                if key not in known_params:
                    # Add to additional parameters list
                    additional_params.append((key, value))

            # If we have additional parameters, add them to the UI
            if additional_params:
                additional_group = QGroupBox("Additional Parameters")
                additional_layout = QVBoxLayout(additional_group)

                # Create a table for additional parameters
                params_table = QTableWidget()
                params_table.setColumnCount(2)
                params_table.setHorizontalHeaderLabels(["Parameter", "Value"])
                params_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
                params_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
                params_table.setRowCount(len(additional_params))

                # Add extraction parameters display
                self.extraction_params_text = QTextEdit()
                self.extraction_params_text.setReadOnly(True)
                self.extraction_params_text.setStyleSheet("color: black; background-color: white;")
                self.extraction_params_text.setMinimumHeight(200)

                # Improve table styling
                params_table.setAlternatingRowColors(True)
                params_table.setShowGrid(True)
                params_table.verticalHeader().setVisible(False)
                params_table.setStyleSheet("""
                    QTableWidget {
                        border: 1px solid #ddd;
                        gridline-color: #ddd;
                        background-color: white;
                        alternate-background-color: #f9f9f9;
                    }
                    QTableWidget::item {
                        padding: 5px;
                        color: black;
                    }
                    QHeaderView::section {
                        background-color: #f0f0f0;
                        padding: 5px;
                        border: none;
                        border-bottom: 1px solid #ddd;
                        font-weight: bold;
                        color: #333;
                    }
                """)

                for i, (key, value) in enumerate(additional_params):
                    # Parameter name
                    key_item = QTableWidgetItem(key)
                    params_table.setItem(i, 0, key_item)

                    # Parameter value
                    if isinstance(value, bool):
                        value_item = QTableWidgetItem("Yes" if value else "No")
                    elif isinstance(value, (dict, list)):
                        # Convert complex values to a readable string format
                        import json
                        try:
                            # Format the JSON with indentation for better readability
                            formatted_value = json.dumps(value, indent=2)
                            value_item = QTableWidgetItem(formatted_value)
                        except:
                            # Fallback if JSON conversion fails
                            value_item = QTableWidgetItem(str(value))
                    else:
                        value_item = QTableWidgetItem(str(value))
                    # Set text alignment and word wrap for better readability
                    value_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    params_table.setItem(i, 1, value_item)

                    # Adjust row height for complex values
                    if isinstance(value, (dict, list)):
                        # Set a taller row height for complex values
                        params_table.setRowHeight(i, 80)
                    else:
                        # Standard row height for simple values
                        params_table.setRowHeight(i, 30)

                # Set word wrap mode for the value column
                params_table.setWordWrap(True)

                additional_layout.addWidget(params_table)
                main_layout.addWidget(additional_group)

            # Add extraction parameters display
            self.extraction_params_text = QTextEdit()
            self.extraction_params_text.setReadOnly(True)
            self.extraction_params_text.setStyleSheet("color: black; background-color: white;")
            self.extraction_params_text.setMinimumHeight(200)

            # Add extraction parameters group
            extraction_params_group = QGroupBox("Extraction Parameters")
            extraction_params_layout = QVBoxLayout(extraction_params_group)
            extraction_params_layout.addWidget(self.extraction_params_text)
            main_layout.addWidget(extraction_params_group)

            # Display extraction parameters
            self.display_extraction_parameters()

        # Add a debug section to show the actual config structure
        debug_btn = QPushButton("Show Raw Config")
        debug_btn.setToolTip("Show the raw configuration structure for debugging")
        debug_btn.clicked.connect(lambda: self.show_raw_config(config))
        main_layout.addWidget(debug_btn)

        main_layout.addStretch()

        return tab

    def create_validation_tab(self):
        """Create the YAML template tab for invoice2data templates"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(20)

        # Add explanation
        explanation = QLabel("Define a YAML template for invoice2data extraction.")
        explanation.setWordWrap(True)
        explanation.setStyleSheet("color: #666;")
        layout.addWidget(explanation)

        # Add more detailed explanation
        details = QLabel(
            "This template defines the structure for invoice2data to extract data from invoices. "
            "It includes issuer, fields, keywords, and options for the extraction process. "
            "YAML format is preferred for invoice2data templates."
        )
        details.setWordWrap(True)
        details.setStyleSheet("color: #666; font-style: italic; margin-bottom: 10px;")
        layout.addWidget(details)

        # Create YAML template editor
        self.json_template_editor = QTextEdit()  # Keep the same variable name for compatibility
        self.json_template_editor.setFont(QFont("Courier New", 10))  # Use monospace font for better YAML editing
        self.json_template_editor.setLineWrapMode(QTextEdit.NoWrap)  # Disable line wrapping for better YAML editing
        layout.addWidget(self.json_template_editor)

        # Load existing JSON template if available
        print(f"\n[DEBUG] Loading JSON template in create_validation_tab")
        print(f"[DEBUG] Has template_data: {hasattr(self, 'template_data') and self.template_data is not None}")
        if hasattr(self, 'template_data') and self.template_data:
            print(f"[DEBUG] template_data keys: {list(self.template_data.keys())}")
            print(f"[DEBUG] 'json_template' in template_data: {'json_template' in self.template_data}")

        if hasattr(self, 'template_data') and self.template_data and 'json_template' in self.template_data:
            json_template = self.template_data['json_template']
            print(f"\n[DEBUG] JSON template found in template data: {json_template is not None}")
            if json_template:
                print(f"[DEBUG] JSON template type: {type(json_template)}")
                if isinstance(json_template, dict):
                    print(f"[DEBUG] JSON template keys: {list(json_template.keys())}")
                    # Format as YAML (preferred)
                    try:
                        import yaml
                        formatted_yaml = yaml.dump(json_template, default_flow_style=False, allow_unicode=True)
                        print(f"[DEBUG] Setting YAML template text (first 200 chars): {formatted_yaml[:200]}...")
                        self.json_template_editor.setText(formatted_yaml)
                    except Exception as e:
                        # Fallback to JSON if YAML formatting fails
                        print(f"[DEBUG] Failed to format as YAML, using JSON: {str(e)}")
                        formatted_json = json.dumps(json_template, indent=2)
                        print(f"[DEBUG] Setting JSON template text (first 200 chars): {formatted_json[:200]}...")
                        self.json_template_editor.setText(formatted_json)
                else:
                    print(f"[DEBUG] JSON template is not a dictionary: {json_template}")
                    # Set default invoice_extractor template structure
                    default_template = {
                        "issuer": "Company Name",
                        "fields": {
                            "invoice_number": "1",
                            "date": "1",
                            "amount": "1"
                        },
                        "keywords": ["Company Name"],
                        "options": {
                            "currency": "INR",
                            "languages": ["en"],
                            "decimal_separator": ".",
                            "replace": []
                        }
                    }

                    # Format as YAML
                    try:
                        import yaml
                        yaml_text = yaml.dump(default_template, default_flow_style=False, allow_unicode=True)
                        self.json_template_editor.setText(yaml_text)
                    except Exception as e:
                        # Fallback to JSON if YAML formatting fails
                        print(f"[DEBUG] Failed to format as YAML, using JSON: {str(e)}")
                        self.json_template_editor.setText(json.dumps(default_template, indent=2))
            else:
                print(f"[DEBUG] JSON template is None or empty, using default template")
                # Set default invoice_extractor template structure
                default_template = {
                    "issuer": "Company Name",
                    "fields": {
                        "invoice_number": "1",
                        "date": "1",
                        "amount": "1"
                    },
                    "keywords": ["Company Name"],
                    "options": {
                        "currency": "INR",
                        "languages": ["en"],
                        "decimal_separator": ".",
                        "replace": []
                    }
                }

                # Format as YAML
                try:
                    import yaml
                    yaml_text = yaml.dump(default_template, default_flow_style=False, allow_unicode=True)
                    self.json_template_editor.setText(yaml_text)
                except Exception as e:
                    # Fallback to JSON if YAML formatting fails
                    print(f"[DEBUG] Failed to format as YAML, using JSON: {str(e)}")
                    self.json_template_editor.setText(json.dumps(default_template, indent=2))
        else:
            print(f"[DEBUG] No JSON template found in template_data, using default template")
            # Use factory to create default template
            default_template = TemplateFactory.create_default_invoice_template("Company Name")

            # Format as YAML
            try:
                import yaml
                yaml_text = yaml.dump(default_template, default_flow_style=False, allow_unicode=True)
                self.json_template_editor.setText(yaml_text)
            except Exception as e:
                # Fallback to JSON if YAML formatting fails
                print(f"[DEBUG] Failed to format as YAML, using JSON: {str(e)}")
                self.json_template_editor.setText(json.dumps(default_template, indent=2))



        # Buttons for validation and reset
        button_layout = QHBoxLayout()

        validate_json_btn = QPushButton("Validate YAML")
        validate_json_btn.clicked.connect(self.validate_json_template)
        validate_json_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3158D3;
            }
        """)

        reset_template_btn = QPushButton("Reset Template")
        reset_template_btn.clicked.connect(self.reset_json_template)
        reset_template_btn.setStyleSheet("""
            QPushButton {
                background-color: #EF4444;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #DC2626;
            }
        """)

        button_layout.addWidget(validate_json_btn)
        button_layout.addWidget(reset_template_btn)
        button_layout.addStretch()

        layout.addLayout(button_layout)

        return tab

    def create_mapping_tab(self):
        """Create the mapping configuration tab"""
        tab = QWidget()
        main_layout = QVBoxLayout(tab)
        main_layout.setSpacing(16)

        # Add explanation
        explanation = QLabel("Configure how template regions are applied to PDF documents with varying page counts.")
        explanation.setWordWrap(True)
        explanation.setStyleSheet("color: #666; margin-bottom: 10px;")
        main_layout.addWidget(explanation)

        # Mapping approach selection
        approach_group = QGroupBox("Mapping Approach")
        approach_layout = QVBoxLayout(approach_group)

        self.mapping_approach_combo = QComboBox()
        self.mapping_approach_combo.addItems(["Page-wise Mapping", "Region-wise Mapping"])

        # Set current selection based on mapping config
        current_approach = self.mapping_config.get('approach', 'page_wise')
        if current_approach == 'region_wise':
            self.mapping_approach_combo.setCurrentText("Region-wise Mapping")
        else:
            self.mapping_approach_combo.setCurrentText("Page-wise Mapping")

        self.mapping_approach_combo.currentTextChanged.connect(self.on_mapping_approach_changed)
        approach_layout.addWidget(self.mapping_approach_combo)

        # Add approach descriptions
        page_wise_desc = QLabel("Maps entire template pages to PDF pages (e.g., template page 1 → PDF page 1)")
        page_wise_desc.setStyleSheet("color: #666; font-style: italic; margin-left: 20px;")
        approach_layout.addWidget(page_wise_desc)

        region_wise_desc = QLabel("Maps specific regions independently across pages (e.g., header from page 1, items from all pages)")
        region_wise_desc.setStyleSheet("color: #666; font-style: italic; margin-left: 20px;")
        approach_layout.addWidget(region_wise_desc)

        main_layout.addWidget(approach_group)

        # Create stacked widget for different mapping configurations
        self.mapping_stack = QStackedWidget()

        # Page-wise mapping configuration
        self.page_wise_widget = self.create_page_wise_mapping_widget()
        self.mapping_stack.addWidget(self.page_wise_widget)

        # Region-wise mapping configuration
        self.region_wise_widget = self.create_region_wise_mapping_widget()
        self.mapping_stack.addWidget(self.region_wise_widget)

        # Set initial stack based on current approach
        if current_approach == 'region_wise':
            self.mapping_stack.setCurrentWidget(self.region_wise_widget)
        else:
            self.mapping_stack.setCurrentWidget(self.page_wise_widget)

        main_layout.addWidget(self.mapping_stack)

        # Preview panel
        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout(preview_group)

        preview_desc = QLabel("Preview how mapping settings apply to documents with different page counts:")
        preview_desc.setStyleSheet("color: #666;")
        preview_layout.addWidget(preview_desc)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(200)
        self.preview_text.setStyleSheet("color: black; background-color: #f9f9f9; border: 1px solid #ddd;")
        preview_layout.addWidget(self.preview_text)

        main_layout.addWidget(preview_group)

        # Update preview
        self.update_mapping_preview()

        main_layout.addStretch()
        return tab

    def create_page_wise_mapping_widget(self):
        """Create the page-wise mapping configuration widget"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Page mapping configuration
        config_group = QGroupBox("Page Mapping Configuration")
        config_layout = QFormLayout(config_group)

        # First page mapping
        self.first_page_spin = QSpinBox()
        self.first_page_spin.setMinimum(1)
        self.first_page_spin.setMaximum(self.page_count)
        self.first_page_spin.setValue(self.mapping_config.get('page_wise', {}).get('first_page', 1))
        self.first_page_spin.valueChanged.connect(self.update_mapping_preview)
        config_layout.addRow("First Page (Template Page):", self.first_page_spin)

        # Middle pages mapping
        self.middle_pages_combo = QComboBox()
        self.middle_pages_combo.addItems(["Sequential", "Repeat Last", "Repeat First"])

        middle_pages_setting = self.mapping_config.get('page_wise', {}).get('middle_pages', 'sequential')
        if middle_pages_setting == 'repeat_last':
            self.middle_pages_combo.setCurrentText("Repeat Last")
        elif middle_pages_setting == 'repeat_first':
            self.middle_pages_combo.setCurrentText("Repeat First")
        else:
            self.middle_pages_combo.setCurrentText("Sequential")

        self.middle_pages_combo.currentTextChanged.connect(self.update_mapping_preview)
        config_layout.addRow("Middle Pages (2 to n-1):", self.middle_pages_combo)

        # Last page mapping
        self.last_page_combo = QComboBox()
        self.last_page_combo.addItems(["Last Template Page", "Specific Page", "Same as First"])

        last_page_setting = self.mapping_config.get('page_wise', {}).get('last_page', 'last_template_page')
        if last_page_setting == 'same_as_first':
            self.last_page_combo.setCurrentText("Same as First")
        elif isinstance(last_page_setting, int):
            self.last_page_combo.setCurrentText("Specific Page")
        else:
            self.last_page_combo.setCurrentText("Last Template Page")

        self.last_page_combo.currentTextChanged.connect(self.on_last_page_option_changed)
        config_layout.addRow("Last Page (n):", self.last_page_combo)

        # Specific last page spinner (only shown when "Specific Page" is selected)
        self.last_page_spin = QSpinBox()
        self.last_page_spin.setMinimum(1)
        self.last_page_spin.setMaximum(self.page_count)
        if isinstance(last_page_setting, int):
            self.last_page_spin.setValue(last_page_setting)
        else:
            self.last_page_spin.setValue(self.page_count)
        self.last_page_spin.valueChanged.connect(self.update_mapping_preview)
        config_layout.addRow("Specific Last Page:", self.last_page_spin)

        # Initially hide the specific page spinner
        self.last_page_spin.setVisible(isinstance(last_page_setting, int))
        config_layout.labelForField(self.last_page_spin).setVisible(isinstance(last_page_setting, int))

        layout.addWidget(config_group)

        # Visual representation
        visual_group = QGroupBox("Visual Representation")
        visual_layout = QVBoxLayout(visual_group)

        visual_desc = QLabel("Template pages will be mapped to PDF pages according to the configuration above.")
        visual_desc.setStyleSheet("color: #666;")
        visual_layout.addWidget(visual_desc)

        layout.addWidget(visual_group)
        layout.addStretch()

        return widget

    def create_region_wise_mapping_widget(self):
        """Create the region-wise mapping configuration widget"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Region mapping configuration
        config_group = QGroupBox("Region Mapping Configuration")
        config_layout = QFormLayout(config_group)

        region_wise_config = self.mapping_config.get('region_wise', {})

        # Header region mapping
        self.header_source_input = QLineEdit()
        self.header_source_input.setPlaceholderText("1")
        header_source = region_wise_config.get('header', {}).get('source_page', '1')
        self.header_source_input.setText(str(header_source))
        self.header_source_input.setToolTip(
            "Specify pages for header regions:\n"
            "• Single page: 1\n"
            "• Multiple pages: 1,3,5\n"
            "• Page ranges: 2-4\n"
            "• Combined: 1,3-5,7\n"
            "• Last page: n or last\n"
            "• Relative: n-1, n-2\n"
            "• Mixed: 1,3-5,n-1,n"
        )
        self.header_source_input.textChanged.connect(self.on_region_input_changed)
        config_layout.addRow("Header Regions Source:", self.header_source_input)

        # Header validation label
        self.header_validation_label = QLabel()
        self.header_validation_label.setStyleSheet("color: #666; font-size: 11px; margin-left: 20px;")
        config_layout.addRow("", self.header_validation_label)

        # Items region mapping
        self.items_source_input = QLineEdit()
        self.items_source_input.setPlaceholderText("1,3-5,n")
        items_source = region_wise_config.get('items', {}).get('source_page', '1-n')
        self.items_source_input.setText(str(items_source))
        self.items_source_input.setToolTip(
            "Specify pages for items regions:\n"
            "• All pages: 1-n\n"
            "• Multiple pages: 1,3,5\n"
            "• Page ranges: 2-4\n"
            "• Combined: 1,3-5,7\n"
            "• Last page: n or last\n"
            "• Relative: n-1, n-2\n"
            "• Mixed: 1,3-5,n-1,n"
        )
        self.items_source_input.textChanged.connect(self.on_region_input_changed)
        config_layout.addRow("Items Regions Source:", self.items_source_input)

        # Items validation label
        self.items_validation_label = QLabel()
        self.items_validation_label.setStyleSheet("color: #666; font-size: 11px; margin-left: 20px;")
        config_layout.addRow("", self.items_validation_label)

        # Summary region mapping
        self.summary_source_input = QLineEdit()
        self.summary_source_input.setPlaceholderText("n")
        summary_source = region_wise_config.get('summary', {}).get('source_page', 'n')
        self.summary_source_input.setText(str(summary_source))
        self.summary_source_input.setToolTip(
            "Specify pages for summary regions:\n"
            "• Last page: n or last\n"
            "• Multiple pages: 1,3,5\n"
            "• Page ranges: 2-4\n"
            "• Combined: 1,3-5,7\n"
            "• Relative: n-1, n-2\n"
            "• Mixed: 1,3-5,n-1,n"
        )
        self.summary_source_input.textChanged.connect(self.on_region_input_changed)
        config_layout.addRow("Summary Regions Source:", self.summary_source_input)

        # Summary validation label
        self.summary_validation_label = QLabel()
        self.summary_validation_label.setStyleSheet("color: #666; font-size: 11px; margin-left: 20px;")
        config_layout.addRow("", self.summary_validation_label)

        layout.addWidget(config_group)

        # Syntax help
        syntax_group = QGroupBox("Syntax Reference")
        syntax_layout = QVBoxLayout(syntax_group)

        syntax_text = QLabel(
            "<b>Supported Syntax:</b><br>"
            "• <b>Single page:</b> 1<br>"
            "• <b>Multiple pages:</b> 1,3,5<br>"
            "• <b>Page ranges:</b> 2-4 (pages 2, 3, 4)<br>"
            "• <b>Combined:</b> 1,3-5,7 (pages 1, 3, 4, 5, 7)<br>"
            "• <b>Last page:</b> n or last<br>"
            "• <b>Relative pages:</b> n-1 (second-to-last), n-2 (third-to-last)<br>"
            "• <b>All pages:</b> 1-n<br>"
            "• <b>Complex example:</b> 1,3-5,n-1,n"
        )
        syntax_text.setWordWrap(True)
        syntax_text.setStyleSheet("color: #666; font-size: 11px;")
        syntax_layout.addWidget(syntax_text)

        layout.addWidget(syntax_group)

        # Initialize validation
        self.on_region_input_changed()

        layout.addStretch()
        return widget

    def validate_json_template(self):
        """Validate the invoice2data YAML template"""
        try:
            # Get the template text from the editor
            template_text = self.json_template_editor.toPlainText()
            if not template_text.strip():
                QMessageBox.warning(self, "Empty Template", "The template is empty. Please enter a valid YAML template.")
                return

            # Try to parse as YAML first
            import yaml
            try:
                template = yaml.safe_load(template_text)
                is_yaml = True
                print(f"[DEBUG] Successfully parsed template as YAML")
            except yaml.YAMLError as yaml_err:
                # If YAML parsing fails, try JSON as fallback
                try:
                    template = json.loads(template_text)
                    is_yaml = False
                    print(f"[DEBUG] YAML parsing failed, but JSON parsing succeeded")
                except json.JSONDecodeError:
                    # Both YAML and JSON parsing failed
                    QMessageBox.critical(self, "Invalid Template",
                                        f"The template contains invalid YAML/JSON syntax: {str(yaml_err)}")
                    return

            # Check if the template has the required structure
            if not isinstance(template, dict):
                QMessageBox.warning(self, "Invalid Template", "The template must be a YAML/JSON object.")
                return

            # Check for required fields in the invoice2data template
            required_fields = ['issuer', 'fields']
            missing_fields = [field for field in required_fields if field not in template]
            if missing_fields:
                QMessageBox.warning(self, "Invalid Template",
                                   f"The template is missing required fields: {', '.join(missing_fields)}")
                return

            # Check that fields contains at least the required fields
            required_invoice_fields = ['invoice_number', 'date', 'amount']
            if 'fields' in template:
                missing_invoice_fields = [field for field in required_invoice_fields if field not in template['fields']]
                if missing_invoice_fields:
                    QMessageBox.warning(self, "Invalid Template",
                                       f"The 'fields' section is missing required fields: {', '.join(missing_invoice_fields)}")
                    return

            # Format the template with proper indentation (prefer YAML)
            try:
                # Format as YAML (preferred)
                formatted_template = yaml.dump(template, default_flow_style=False, allow_unicode=True)
                self.json_template_editor.setText(formatted_template)
                format_type = "YAML"
            except Exception as yaml_format_err:
                # Fallback to JSON if YAML formatting fails
                formatted_template = json.dumps(template, indent=2, ensure_ascii=False)
                self.json_template_editor.setText(formatted_template)
                format_type = "JSON"
                print(f"[DEBUG] YAML formatting failed, using JSON format: {str(yaml_format_err)}")

            QMessageBox.information(self, "Valid Template", f"The invoice2data template is valid and has been formatted as {format_type}.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while validating the template: {str(e)}")

    def reset_json_template(self):
        """Reset the template to the default structure in YAML format"""
        reply = QMessageBox.question(
            self,
            "Reset Template",
            "Are you sure you want to reset the template to the default structure? This will discard all your changes.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # Use factory to create default template
            default_template = TemplateFactory.create_default_invoice_template("Company Name")

            # Format as YAML
            try:
                import yaml
                yaml_text = yaml.dump(default_template, default_flow_style=False, allow_unicode=True)
                self.json_template_editor.setText(yaml_text)
            except Exception as e:
                # Fallback to JSON if YAML formatting fails
                print(f"[DEBUG] Failed to format as YAML, using JSON: {str(e)}")
                self.json_template_editor.setText(json.dumps(default_template, indent=2))

    # Removed add_validation_rule and remove_validation_rule functions as they are no longer needed

    def on_template_type_changed(self, template_type):
        """Handle template type change"""
        is_multi_page = template_type == "Multi-page"
        self.update_page_count_visibility()
        if is_multi_page:
            self.page_count = max(1, self.page_count)
            self.page_count_spin.setValue(self.page_count)
        else:
            self.page_count = 1
            self.page_count_spin.setValue(1)

    def update_page_count_visibility(self):
        """Show/hide page count based on template type"""
        is_multi_page = self.type_combo.currentText() == "Multi-page"
        self.page_count_container.setVisible(is_multi_page)

    def on_page_count_changed(self, value):
        """Handle page count change"""
        self.page_count = value
        if self.current_page >= value:
            self.current_page = value - 1
        self.update_page_navigation()
        # Update mapping configuration spinboxes
        if hasattr(self, 'first_page_spin'):
            self.first_page_spin.setMaximum(value)
        if hasattr(self, 'last_page_spin'):
            self.last_page_spin.setMaximum(value)
        self.update_mapping_preview()

    def on_mapping_approach_changed(self, approach_text):
        """Handle mapping approach change"""
        if approach_text == "Region-wise Mapping":
            self.mapping_stack.setCurrentWidget(self.region_wise_widget)
            self.mapping_config['approach'] = 'region_wise'
        else:
            self.mapping_stack.setCurrentWidget(self.page_wise_widget)
            self.mapping_config['approach'] = 'page_wise'
        self.update_mapping_preview()

    def on_last_page_option_changed(self, option_text):
        """Handle last page option change"""
        show_spinner = option_text == "Specific Page"
        self.last_page_spin.setVisible(show_spinner)

        # Find the form layout and show/hide the label
        form_layout = self.last_page_spin.parent().layout()
        if isinstance(form_layout, QFormLayout):
            label = form_layout.labelForField(self.last_page_spin)
            if label:
                label.setVisible(show_spinner)

        self.update_mapping_preview()

    def on_region_input_changed(self):
        """Handle region input field changes with validation"""
        try:
            # Validate header input
            if hasattr(self, 'header_source_input') and hasattr(self, 'header_validation_label'):
                header_text = self.header_source_input.text().strip()
                header_pages, header_error = self.parse_page_expression(header_text, 5)  # Test with 5 pages
                if header_error:
                    self.header_validation_label.setText(f"❌ {header_error}")
                    self.header_validation_label.setStyleSheet("color: #d32f2f; font-size: 11px; margin-left: 20px;")
                else:
                    self.header_validation_label.setText(f"✓ Pages: {sorted(header_pages)}")
                    self.header_validation_label.setStyleSheet("color: #388e3c; font-size: 11px; margin-left: 20px;")

            # Validate items input
            if hasattr(self, 'items_source_input') and hasattr(self, 'items_validation_label'):
                items_text = self.items_source_input.text().strip()
                items_pages, items_error = self.parse_page_expression(items_text, 5)  # Test with 5 pages
                if items_error:
                    self.items_validation_label.setText(f"❌ {items_error}")
                    self.items_validation_label.setStyleSheet("color: #d32f2f; font-size: 11px; margin-left: 20px;")
                else:
                    self.items_validation_label.setText(f"✓ Pages: {sorted(items_pages)}")
                    self.items_validation_label.setStyleSheet("color: #388e3c; font-size: 11px; margin-left: 20px;")

            # Validate summary input
            if hasattr(self, 'summary_source_input') and hasattr(self, 'summary_validation_label'):
                summary_text = self.summary_source_input.text().strip()
                summary_pages, summary_error = self.parse_page_expression(summary_text, 5)  # Test with 5 pages
                if summary_error:
                    self.summary_validation_label.setText(f"❌ {summary_error}")
                    self.summary_validation_label.setStyleSheet("color: #d32f2f; font-size: 11px; margin-left: 20px;")
                else:
                    self.summary_validation_label.setText(f"✓ Pages: {sorted(summary_pages)}")
                    self.summary_validation_label.setStyleSheet("color: #388e3c; font-size: 11px; margin-left: 20px;")

            # Update preview
            self.update_mapping_preview()

        except Exception as e:
            print(f"Error in region input validation: {e}")

    def parse_page_expression(self, expression, total_pages):
        """
        Parse a page expression and return the list of pages and any error message.

        Args:
            expression (str): Page expression like "1,3-5,n-1,n"
            total_pages (int): Total number of pages in the document

        Returns:
            tuple: (list of page numbers, error message or None)
        """
        if not expression.strip():
            return [], "Empty expression"

        try:
            pages = set()
            parts = [part.strip() for part in expression.split(',')]

            for part in parts:
                if not part:
                    continue

                # Handle ranges (e.g., "2-4", "1-n", "n-2-n")
                if '-' in part and not part.startswith('n-') and not part == 'n':
                    # Split on dash, but be careful with negative relative references
                    if part.count('-') == 1:
                        start_str, end_str = part.split('-', 1)
                    else:
                        # Handle cases like "n-2" which should not be split
                        if part.startswith('n-') and part.count('-') == 1:
                            # This is a relative reference like "n-2", not a range
                            page_num = self.resolve_page_reference(part, total_pages)
                            if page_num is None:
                                return [], f"Invalid page reference: {part}"
                            if 1 <= page_num <= total_pages:
                                pages.add(page_num)
                            continue
                        else:
                            return [], f"Invalid range format: {part}"

                    start_page = self.resolve_page_reference(start_str, total_pages)
                    end_page = self.resolve_page_reference(end_str, total_pages)

                    if start_page is None:
                        return [], f"Invalid start page: {start_str}"
                    if end_page is None:
                        return [], f"Invalid end page: {end_str}"

                    if start_page > end_page:
                        return [], f"Invalid range: {start_page} > {end_page}"

                    for page in range(start_page, end_page + 1):
                        if 1 <= page <= total_pages:
                            pages.add(page)
                else:
                    # Handle single page references
                    page_num = self.resolve_page_reference(part, total_pages)
                    if page_num is None:
                        return [], f"Invalid page reference: {part}"
                    if 1 <= page_num <= total_pages:
                        pages.add(page_num)

            return list(pages), None

        except Exception as e:
            return [], f"Parse error: {str(e)}"

    def resolve_page_reference(self, ref, total_pages):
        """
        Resolve a single page reference to an actual page number.

        Args:
            ref (str): Page reference like "1", "n", "n-1", "last"
            total_pages (int): Total number of pages

        Returns:
            int or None: Resolved page number, or None if invalid
        """
        ref = ref.strip().lower()

        if not ref:
            return None

        # Handle numeric references
        if ref.isdigit():
            return int(ref)

        # Handle "last" keyword
        if ref == 'last':
            return total_pages

        # Handle "n" (last page)
        if ref == 'n':
            return total_pages

        # Handle relative references like "n-1", "n-2"
        if ref.startswith('n-'):
            try:
                offset_str = ref[2:]  # Remove "n-"
                if offset_str.isdigit():
                    offset = int(offset_str)
                    result = total_pages - offset
                    return result if result >= 1 else None
                else:
                    return None
            except:
                return None

        return None

    def update_mapping_preview(self):
        """Update the mapping preview text"""
        try:
            preview_text = "Mapping Preview:\n\n"

            current_approach = self.mapping_config.get('approach', 'page_wise')

            # Test with different document page counts
            test_page_counts = [1, 3, 5]

            for doc_pages in test_page_counts:
                preview_text += f"📄 {doc_pages}-page document:\n"

                if current_approach == 'page_wise':
                    preview_text += self.generate_page_wise_preview(doc_pages)
                else:
                    preview_text += self.generate_region_wise_preview(doc_pages)

                preview_text += "\n"

            self.preview_text.setText(preview_text)
        except Exception as e:
            print(f"Error updating mapping preview: {e}")
            self.preview_text.setText("Error generating preview")

    def generate_page_wise_preview(self, doc_pages):
        """Generate preview text for page-wise mapping"""
        preview = ""

        # Get current settings
        first_page = getattr(self, 'first_page_spin', None)
        first_page_val = first_page.value() if first_page else 1

        middle_pages = getattr(self, 'middle_pages_combo', None)
        middle_pages_val = middle_pages.currentText() if middle_pages else "Sequential"

        last_page = getattr(self, 'last_page_combo', None)
        last_page_val = last_page.currentText() if last_page else "Last Template Page"

        last_page_spin = getattr(self, 'last_page_spin', None)
        last_page_spin_val = last_page_spin.value() if last_page_spin else self.page_count

        for pdf_page in range(1, doc_pages + 1):
            if pdf_page == 1:
                template_page = first_page_val
            elif pdf_page == doc_pages and doc_pages > 1:
                if last_page_val == "Last Template Page":
                    template_page = self.page_count
                elif last_page_val == "Specific Page":
                    template_page = last_page_spin_val
                else:  # Same as First
                    template_page = first_page_val
            else:  # Middle pages
                if middle_pages_val == "Sequential":
                    template_page = min(pdf_page, self.page_count)
                elif middle_pages_val == "Repeat Last":
                    template_page = self.page_count
                else:  # Repeat First
                    template_page = first_page_val

            preview += f"  Page {pdf_page} → Template Page {template_page}\n"

        return preview

    def generate_region_wise_preview(self, doc_pages):
        """Generate preview text for region-wise mapping"""
        preview = ""

        # Get current settings from text inputs
        header_input = getattr(self, 'header_source_input', None)
        header_text = header_input.text().strip() if header_input else "1"

        items_input = getattr(self, 'items_source_input', None)
        items_text = items_input.text().strip() if items_input else "1-n"

        summary_input = getattr(self, 'summary_source_input', None)
        summary_text = summary_input.text().strip() if summary_input else "n"

        # Parse page expressions
        header_pages, header_error = self.parse_page_expression(header_text, doc_pages)
        items_pages, items_error = self.parse_page_expression(items_text, doc_pages)
        summary_pages, summary_error = self.parse_page_expression(summary_text, doc_pages)

        # Display results
        if header_error:
            preview += f"  Header regions: ❌ {header_error}\n"
        else:
            preview += f"  Header regions: {header_text} → Pages {sorted(header_pages)}\n"

        if items_error:
            preview += f"  Items regions: ❌ {items_error}\n"
        else:
            preview += f"  Items regions: {items_text} → Pages {sorted(items_pages)}\n"

        if summary_error:
            preview += f"  Summary regions: ❌ {summary_error}\n"
        else:
            preview += f"  Summary regions: {summary_text} → Pages {sorted(summary_pages)}\n"

        return preview

    def prev_page(self):
        """Navigate to previous page"""
        if self.current_page > 0:
            self.current_page -= 1
            self.update_page_navigation()
            self.load_regions_for_current_page()
            self.load_column_lines_for_current_page()
            # Refresh the config tab to show page-specific config
            self.refresh_config_tab()

    def next_page(self):
        """Navigate to next page"""
        if self.current_page < self.page_count - 1:
            self.current_page += 1
            self.update_page_navigation()
            self.load_regions_for_current_page()
            self.load_column_lines_for_current_page()
            # Refresh the config tab to show page-specific config
            self.refresh_config_tab()

    def update_page_navigation(self):
        """Update page navigation UI"""
        if hasattr(self, 'page_label'):
            self.page_label.setText(f"Page {self.current_page + 1} of {self.page_count}")
        if hasattr(self, 'page_label_cols'):
            self.page_label_cols.setText(f"Page {self.current_page + 1} of {self.page_count}")
        if hasattr(self, 'page_label_config'):
            self.page_label_config.setText(f"Page {self.current_page + 1} of {self.page_count}")

        # Update button states
        if hasattr(self, 'prev_page_btn'):
            self.prev_page_btn.setEnabled(self.current_page > 0)
        if hasattr(self, 'next_page_btn'):
            self.next_page_btn.setEnabled(self.current_page < self.page_count - 1)
        if hasattr(self, 'prev_page_btn_cols'):
            self.prev_page_btn_cols.setEnabled(self.current_page > 0)
        if hasattr(self, 'next_page_btn_cols'):
            self.next_page_btn_cols.setEnabled(self.current_page < self.page_count - 1)
        if hasattr(self, 'prev_page_btn_config'):
            self.prev_page_btn_config.setEnabled(self.current_page > 0)
        if hasattr(self, 'next_page_btn_config'):
            self.next_page_btn_config.setEnabled(self.current_page < self.page_count - 1)

        # Update config fields if page-specific config is enabled
        if (hasattr(self, 'page_specific_config') and
            self.page_specific_config.isChecked() and
            'page_configs' in self.template_data):
            pass  # This is handled by refresh_config_tab

    def refresh_config_tab(self):
        """Refresh the configuration tab to show page-specific config"""
        try:
            # Find the tab widget
            tab_widget = None
            for child in self.children():
                if isinstance(child, QTabWidget):
                    tab_widget = child
                    break

            if tab_widget:
                # Find the config tab index
                config_tab_index = -1
                for i in range(tab_widget.count()):
                    if tab_widget.tabText(i) == "Configuration":
                        config_tab_index = i
                        break

                if config_tab_index >= 0:
                    # Get the current tab index
                    current_tab_index = tab_widget.currentIndex()

                    # Remove the old config tab
                    tab_widget.removeTab(config_tab_index)

                    # Create a new config tab
                    new_config_tab = self.create_config_tab()

                    # Insert the new config tab at the same position
                    tab_widget.insertTab(config_tab_index, new_config_tab, "Configuration")

                    # Restore the current tab
                    tab_widget.setCurrentIndex(current_tab_index)

                    print(f"[DEBUG] Refreshed config tab for page {self.current_page + 1}")
                else:
                    print("[DEBUG] Could not find Configuration tab")
            else:
                print("[DEBUG] Could not find tab widget")
        except Exception as e:
            print(f"[ERROR] Failed to refresh config tab: {str(e)}")
            import traceback
            traceback.print_exc()

            page_configs = self.template_data['page_configs']
            if self.current_page < len(page_configs) and page_configs[self.current_page]:
                # Load page-specific config
                if hasattr(self, 'load_page_specific_config'):
                    self.load_page_specific_config(page_configs[self.current_page])

    def load_regions_for_current_page(self):
        """Load regions for the current page"""
        self.regions_table.setRowCount(0)

        # Get regions for current page
        regions = self.get_regions_for_current_page()
        print(f"\n[DEBUG] Loading regions for page {self.current_page + 1}")
        print(f"[DEBUG] Regions data type: {type(regions)}")
        print(f"[DEBUG] Regions keys: {list(regions.keys())}")
        for section, rects in regions.items():
            print(f"[DEBUG] Section '{section}' has {len(rects)} rectangles")
            if rects and len(rects) > 0:
                print(f"[DEBUG] First rect type: {type(rects[0])}")
                print(f"[DEBUG] First rect data: {rects[0]}")

        row = 0
        for section, rects in regions.items():
            for i, rect in enumerate(rects):
                self.regions_table.insertRow(row)
                self.regions_table.setItem(row, 0, QTableWidgetItem(section.title()))
                self.regions_table.setItem(row, 1, QTableWidgetItem(str(i + 1)))

                # Handle different region formats
                if isinstance(rect, QRect):
                    # QRect format (x,y,width,height)
                    x = rect.x()
                    y = rect.y()
                    width = rect.width()
                    height = rect.height()
                elif hasattr(rect, 'drawing_x') and hasattr(rect, 'drawing_y'):
                    # DualCoordinateRegion format - use drawing coordinates for UI display
                    from dual_coordinate_storage import DualCoordinateRegion
                    if isinstance(rect, DualCoordinateRegion):
                        x = rect.drawing_x
                        y = rect.drawing_y
                        width = rect.drawing_width
                        height = rect.drawing_height
                        print(f"[DEBUG] Displaying DualCoordinateRegion {rect.label}: UI({x},{y},{width},{height})")
                    else:
                        print(f"Warning: Unknown DualCoordinateRegion-like object in {section}: {type(rect)}")
                        continue
                elif hasattr(rect, 'rect') and hasattr(rect, 'label'):
                    # StandardRegion format - use UI coordinates for display
                    from standardized_coordinates import StandardRegion
                    if isinstance(rect, StandardRegion):
                        x = rect.rect.x()
                        y = rect.rect.y()
                        width = rect.rect.width()
                        height = rect.rect.height()
                        print(f"[DEBUG] Displaying StandardRegion {rect.label}: UI({x},{y},{width},{height})")
                    else:
                        print(f"Warning: Unknown StandardRegion-like object in {section}: {type(rect)}")
                        continue
                elif isinstance(rect, dict):
                    if 'x' in rect and 'y' in rect and 'width' in rect and 'height' in rect:
                        # Dictionary format (x,y,width,height)
                        x = rect['x']
                        y = rect['y']
                        width = rect['width']
                        height = rect['height']
                    elif 'x1' in rect and 'y1' in rect and 'x2' in rect and 'y2' in rect:
                        # Dictionary format (x1,y1,x2,y2)
                        x = rect['x1']
                        y = rect['y1']
                        width = rect['x2'] - rect['x1']
                        height = rect['y2'] - rect['y1']
                    else:
                        print(f"Warning: Unknown region format in {section}: {rect}")
                        continue
                else:
                    print(f"Warning: Unknown region type in {section}: {type(rect)}")
                    continue

                # Store drawn format (x,y,width,height)
                self.regions_table.setItem(row, 2, QTableWidgetItem(str(x)))
                self.regions_table.setItem(row, 3, QTableWidgetItem(str(y)))
                self.regions_table.setItem(row, 4, QTableWidgetItem(str(width)))
                self.regions_table.setItem(row, 5, QTableWidgetItem(str(height)))

                # Store scaled format (x1,y1,x2,y2)
                x1 = x
                y1 = y
                x2 = x + width
                y2 = y + height
                scaled_text = f"({x1}, {y1}, {x2}, {y2})"
                scaled_item = QTableWidgetItem(scaled_text)
                scaled_item.setToolTip("Format: x1, y1, x2, y2")
                self.regions_table.setItem(row, 6, scaled_item)

                row += 1

    def load_column_lines_for_current_page(self):
        """Load column lines for the current page"""
        try:
            # Clear existing items
            self.columns_table.setRowCount(0)

            # Get current page's column lines
            column_lines = self.get_column_lines_for_current_page()
            print(f"\n[DEBUG] Loading column lines for page {self.current_page + 1}")
            print(f"[DEBUG] Column lines data type: {type(column_lines)}")
            print(f"[DEBUG] Column lines keys: {list(column_lines.keys()) if isinstance(column_lines, dict) else 'Not a dict'}")

            if isinstance(column_lines, dict):
                for section, lines in column_lines.items():
                    print(f"[DEBUG] Section '{section}' has {len(lines)} lines")
                    if lines and len(lines) > 0:
                        print(f"[DEBUG] First line type: {type(lines[0])}")
                        print(f"[DEBUG] First line data: {lines[0]}")

            if not column_lines:
                print("[DEBUG] No column lines found for this page")
                return

            # Add column lines to the table
            for section, lines in column_lines.items():
                if not lines:
                    continue

                for line in lines:
                    try:
                        # Debug output to see the line data structure
                        print(f"\nProcessing line in {section}:")
                        print(f"Line type: {type(line)}")
                        print(f"Line data: {line}")

                        # Handle different column line formats
                        if hasattr(line, 'drawing_start_x') and hasattr(line, 'drawing_start_y'):
                            # DualCoordinateColumnLine format - use drawing coordinates for UI display
                            from dual_coordinate_storage import DualCoordinateColumnLine
                            if isinstance(line, DualCoordinateColumnLine):
                                start_x = line.drawing_start_x
                                start_y = line.drawing_start_y
                                end_x = line.drawing_end_x
                                end_y = line.drawing_end_y
                                region_index = 0  # Default region index
                                print(f"[DEBUG] Displaying DualCoordinateColumnLine {line.label}: UI({start_x},{start_y}) -> ({end_x},{end_y})")
                            else:
                                print(f"Warning: Unknown DualCoordinateColumnLine-like object in {section}: {type(line)}")
                                continue
                        elif isinstance(line, (list, tuple)):
                            if len(line) >= 2:
                                # Old format: [QPoint, QPoint, region_index] or [dict, dict, region_index]
                                start_point = line[0]
                                end_point = line[1]
                                region_index = line[2] if len(line) > 2 else 0

                                # Convert dictionary to QPoint if needed
                                if isinstance(start_point, dict) and 'x' in start_point:
                                    start_point = QPoint(start_point['x'], start_point['y'])
                                if isinstance(end_point, dict) and 'x' in end_point:
                                    end_point = QPoint(end_point['x'], end_point['y'])
                                else:
                                    print(f"Warning: Invalid line format in {section}, skipping")
                                    continue
                        elif isinstance(line, dict):
                            # New format: {'x1': float, 'y1': float, 'x2': float, 'y2': float, 'region_index': int}
                            if 'x1' in line and 'y1' in line and 'x2' in line and 'y2' in line:
                                start_point = QPoint(int(line['x1']), int(line['y1']))
                                end_point = QPoint(int(line['x2']), int(line['y2']))
                                region_index = line.get('region_index', 0)
                            else:
                                print(f"Warning: Invalid dictionary format in {section}, missing coordinates")
                                print(f"Available keys: {line.keys()}")
                                continue
                        else:
                            print(f"Warning: Unknown line format in {section}, skipping")
                            continue

                        # Add row to table
                        row = self.columns_table.rowCount()
                        self.columns_table.insertRow(row)

                        # Add section
                        section_item = QTableWidgetItem(section)
                        section_item.setFlags(section_item.flags() & ~Qt.ItemIsEditable)
                        self.columns_table.setItem(row, 0, section_item)

                        # Add table number (region index + 1)
                        table_item = QTableWidgetItem(str(region_index + 1))
                        table_item.setFlags(table_item.flags() & ~Qt.ItemIsEditable)
                        self.columns_table.setItem(row, 1, table_item)

                        # Handle different coordinate formats for display
                        if hasattr(line, 'drawing_start_x'):
                            # DualCoordinateColumnLine format
                            x_pos_item = QTableWidgetItem(f"{start_x:.1f}")
                            desc = f"Start: ({start_x}, {start_y}) End: ({end_x}, {end_y})"
                        else:
                            # Legacy QPoint format
                            x_pos_item = QTableWidgetItem(f"{start_point.x():.1f}")
                            desc = f"Start: ({start_point.x()}, {start_point.y()}) End: ({end_point.x()}, {end_point.y()})"

                        x_pos_item.setFlags(x_pos_item.flags() & ~Qt.ItemIsEditable)
                        self.columns_table.setItem(row, 2, x_pos_item)

                        # Add description with coordinates
                        desc_item = QTableWidgetItem(desc)
                        desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)
                        self.columns_table.setItem(row, 3, desc_item)

                    except Exception as e:
                        print(f"Error processing line in {section}: {str(e)}")
                        continue

            # Adjust column widths
            self.columns_table.resizeColumnsToContents()

        except Exception as e:
            print(f"Error loading column lines: {str(e)}")
            import traceback
            traceback.print_exc()

    def get_regions_for_current_page(self):
        """Get regions for the current page - prioritize dual coordinate data"""
        if self.template_data.get("template_type") == "multi":
            # For multi-page templates, check dual coordinate data first
            drawing_page_regions = self.template_data.get("drawing_page_regions", [])
            if drawing_page_regions and self.current_page < len(drawing_page_regions):
                print(f"[DEBUG] Using dual coordinate page regions for page {self.current_page + 1}")
                return drawing_page_regions[self.current_page]

            # Fallback to legacy page_regions
            page_regions = self.template_data.get("page_regions", [])
            if self.current_page < len(page_regions):
                print(f"[DEBUG] Using legacy page regions for page {self.current_page + 1}")
                return page_regions[self.current_page]
            return {}
        else:
            # For single-page templates, check dual coordinate data first
            drawing_regions = self.template_data.get("drawing_regions")
            if drawing_regions:
                print(f"[DEBUG] Using dual coordinate regions for single-page template")
                return drawing_regions

            # No dual coordinate data found - return empty
            print(f"[DEBUG] No dual coordinate regions found for single-page template")
            return {'header': [], 'items': [], 'summary': []}

    def get_column_lines_for_current_page(self):
        """Get column lines for the current page - prioritize dual coordinate data"""
        if self.template_data.get("template_type") == "multi":
            # For multi-page templates, check dual coordinate data first
            drawing_page_column_lines = self.template_data.get("drawing_page_column_lines", [])
            if drawing_page_column_lines and self.current_page < len(drawing_page_column_lines):
                print(f"[DEBUG] Using dual coordinate page column lines for page {self.current_page + 1}")
                return drawing_page_column_lines[self.current_page]

            # Fallback to legacy page_column_lines
            page_column_lines = self.template_data.get("page_column_lines", [])
            if self.current_page < len(page_column_lines):
                print(f"[DEBUG] Using legacy page column lines for page {self.current_page + 1}")
                return page_column_lines[self.current_page]
            return {}
        else:
            # For single-page templates, check dual coordinate data first
            drawing_column_lines = self.template_data.get("drawing_column_lines")
            if drawing_column_lines:
                print(f"[DEBUG] Using dual coordinate column lines for single-page template")
                return drawing_column_lines

            # No dual coordinate data found - return empty
            print(f"[DEBUG] No dual coordinate column lines found for single-page template")
            return {'header': [], 'items': [], 'summary': []}

    def get_template_data(self):
        """Get all template data from the dialog"""
        template_data = {}

        # Get general information
        template_data['name'] = self.name_input.text().strip()
        template_data['description'] = self.desc_input.toPlainText().strip()
        template_data['template_type'] = "multi" if self.type_combo.currentText() == "Multi-page" else "single"
        template_data['page_count'] = self.page_count

        print(f"\nCollecting template data for {template_data['name']} (type: {template_data['template_type']})")

        # Get dual coordinate regions data
        from dual_coordinate_storage import DualCoordinateStorage
        storage = DualCoordinateStorage()

        if template_data['template_type'] == "multi":
            # For multi-page templates, collect dual coordinate regions for each page
            print(f"Collecting multi-page dual coordinate regions for {self.page_count} pages")
            drawing_page_regions = []
            extraction_page_regions = []

            for page in range(self.page_count):
                self.current_page = page
                page_dual_regions = self.get_dual_regions_data()

                # Log the regions collected for each page
                region_counts = {section: len(rects) for section, rects in page_dual_regions.items()}
                print(f"- Page {page+1}: collected dual coordinate regions = {region_counts}")

                drawing_page_regions.append(page_dual_regions)
                extraction_page_regions.append(page_dual_regions)  # Same data, different usage

            template_data['drawing_page_regions'] = drawing_page_regions
            template_data['extraction_page_regions'] = extraction_page_regions
            print(f"Multi-page template: collected {len(drawing_page_regions)} page dual coordinate regions")
        else:
            # For single-page templates
            dual_regions = self.get_dual_regions_data()

            # Log the regions collected
            region_counts = {section: len(rects) for section, rects in dual_regions.items()}
            print(f"Single-page template: collected dual coordinate regions = {region_counts}")

            template_data['drawing_regions'] = dual_regions
            template_data['extraction_regions'] = dual_regions  # Same data, different usage

        # Get column lines data
        if template_data['template_type'] == "multi":
            # For multi-page templates, collect column lines for each page
            print(f"Collecting multi-page column lines for {self.page_count} pages")
            page_column_lines = []
            for page in range(self.page_count):
                self.current_page = page
                page_column_line = self.get_column_lines_data()
                # Log the column lines collected for each page
                column_counts = {section: len(lines) for section, lines in page_column_line.items()}
                print(f"- Page {page+1}: collected column lines = {column_counts}")
                page_column_lines.append(page_column_line)
            template_data['page_column_lines'] = page_column_lines

            # Also include an empty 'column_lines' field to satisfy older code
            template_data['column_lines'] = {}
            print(f"Multi-page template: collected {len(page_column_lines)} page_column_lines")
        else:
            # For single-page templates
            column_lines = self.get_column_lines_data()
            template_data['column_lines'] = column_lines
            # Log the column lines collected
            column_counts = {section: len(lines) for section, lines in column_lines.items()}
            print(f"Single-page template: collected column lines = {column_counts}")

            # Initialize page_column_lines as an empty list to satisfy code that might look for it
            template_data['page_column_lines'] = []

        # Get configuration data
        template_data['config'] = self.get_config_data()

        # Get mapping configuration
        template_data['mapping_config'] = self.get_mapping_config_data()

        # Get YAML/JSON template from editor
        try:
            print(f"\n[DEBUG] Getting template from editor in get_template_data")
            template_text = self.json_template_editor.toPlainText()
            print(f"[DEBUG] Template text length: {len(template_text)}")
            print(f"[DEBUG] Template text (first 100 chars): {template_text[:100]}...")

            if template_text.strip():
                # Try to parse as YAML first
                import yaml
                try:
                    template_data_obj = yaml.safe_load(template_text)
                    print(f"[DEBUG] Template parsed successfully as YAML")
                except yaml.YAMLError as yaml_err:
                    # If YAML parsing fails, try JSON as fallback
                    try:
                        template_data_obj = json.loads(template_text)
                        print(f"[DEBUG] YAML parsing failed, but JSON parsing succeeded")
                    except json.JSONDecodeError as json_err:
                        # Both YAML and JSON parsing failed
                        print(f"[WARNING] Invalid template: YAML error: {str(yaml_err)}, JSON error: {str(json_err)}")
                        QMessageBox.warning(
                            self,
                            "Invalid Template",
                            f"The template contains invalid YAML and JSON syntax.\n\nYAML Error: {str(yaml_err)}\n\nJSON Error: {str(json_err)}"
                        )
                        template_data["json_template"] = None
                        return template_data

                print(f"[DEBUG] Template data type: {type(template_data_obj)}")
                if isinstance(template_data_obj, dict):
                    print(f"[DEBUG] Template keys: {list(template_data_obj.keys())}")
                    template_data["json_template"] = template_data_obj
                    print(f"[DEBUG] Added template to template_data")
                else:
                    print(f"[WARNING] Template is not a dictionary: {type(template_data_obj)}")
                    QMessageBox.warning(
                        self,
                        "Invalid Template",
                        f"The template must be a YAML/JSON object, but got {type(template_data_obj).__name__}."
                    )
                    template_data["json_template"] = None
            else:
                print(f"[DEBUG] Template text is empty, setting to None")
                template_data["json_template"] = None
        except Exception as e:
            print(f"[ERROR] Error processing template: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.warning(
                self,
                "Error Processing Template",
                f"An unexpected error occurred while processing the template:\n\n{str(e)}"
            )
            template_data["json_template"] = None

        # For backward compatibility, keep empty validation_rules
        template_data["validation_rules"] = {}

        # Ensure we have valid data before returning
        try:
            self.validate_template_data(template_data)
            print("Template data validation successful")
        except Exception as e:
            print(f"Template data validation failed: {str(e)}")

        return template_data

    def get_regions_data(self):
        """Extract regions data from the regions table"""
        regions = {
            'header': [],
            'items': [],
            'summary': []
        }

        # Iterate through all rows in the regions table
        for row in range(self.regions_table.rowCount()):
            # Get section, which should be in the first column
            section_item = self.regions_table.item(row, 0)
            if not section_item:
                continue

            section = section_item.text().lower()

            # Make sure section is valid
            if section not in regions:
                continue

            # Get coordinates
            try:
                x_item = self.regions_table.item(row, 2)
                y_item = self.regions_table.item(row, 3)
                width_item = self.regions_table.item(row, 4)
                height_item = self.regions_table.item(row, 5)

                if x_item and y_item and width_item and height_item:
                    x = int(x_item.text())
                    y = int(y_item.text())
                    width = int(width_item.text())
                    height = int(height_item.text())

                    # Create dual coordinate region
                    from dual_coordinate_storage import DualCoordinateRegion

                    try:
                        # Generate name for the region
                        name = f"{section[0].upper()}{len(regions[section]) + 1}"

                        # Use default scale factors (these will be updated when template is applied)
                        scale_x, scale_y = 1.0, 1.0
                        page_height = 842.0  # A4 page height in points

                        # Create dual coordinate region
                        dual_region = DualCoordinateRegion.from_ui_input(
                            x, y, width, height, name, scale_x, scale_y, page_height
                        )
                        regions[section].append(dual_region)
                        print(f"Created dual coordinate region with name: {name}")
                    except ValueError as e:
                        print(f"Error creating dual coordinate region: {e}")
                        continue
            except (ValueError, AttributeError) as e:
                print(f"Error parsing region data: {str(e)}")
                continue

        # If no regions were found, preserve the original regions based on template type
        if not any(regions.values()) and hasattr(self, 'template_data'):
            print("No regions found in table, preserving original regions data")

            # Check if this is a multi-page template
            if self.template_data.get("template_type") == "multi":
                # For multi-page templates, get page-specific regions
                page_regions = self.template_data.get("page_regions", [])
                if hasattr(self, 'current_page') and self.current_page < len(page_regions):
                    print(f"Preserving multi-page regions for page {self.current_page}")
                    return page_regions[self.current_page]

            # Otherwise fallback to dual coordinate regions (for single-page templates)
            drawing_regions = self.template_data.get('drawing_regions')
            if drawing_regions:
                print("Preserving single-page dual coordinate regions")
                return drawing_regions
            else:
                print("No regions found - returning empty")
                return {'header': [], 'items': [], 'summary': []}

        return regions

    def get_dual_regions_data(self):
        """Extract dual coordinate regions data from the regions table"""
        from dual_coordinate_storage import DualCoordinateRegion

        regions = {
            'header': [],
            'items': [],
            'summary': []
        }

        # Iterate through all rows in the regions table
        for row in range(self.regions_table.rowCount()):
            # Get section, which should be in the first column
            section_item = self.regions_table.item(row, 0)
            if not section_item:
                continue

            section = section_item.text().lower()

            # Make sure section is valid
            if section not in regions:
                continue

            # Get coordinates
            try:
                x_item = self.regions_table.item(row, 2)
                y_item = self.regions_table.item(row, 3)
                width_item = self.regions_table.item(row, 4)
                height_item = self.regions_table.item(row, 5)

                if x_item and y_item and width_item and height_item:
                    x = int(x_item.text())
                    y = int(y_item.text())
                    width = int(width_item.text())
                    height = int(height_item.text())

                    # Generate name for the region
                    name = f"{section[0].upper()}{len(regions[section]) + 1}"

                    # Use default scale factors (these will be updated when template is applied)
                    scale_x, scale_y = 1.0, 1.0
                    page_height = 842.0  # A4 page height in points

                    # Create dual coordinate region
                    dual_region = DualCoordinateRegion.from_ui_input(
                        x, y, width, height, name, scale_x, scale_y, page_height
                    )
                    regions[section].append(dual_region)

            except (ValueError, AttributeError) as e:
                print(f"Error parsing dual coordinate region data: {str(e)}")
                continue

        return regions

    def get_column_lines_data(self):
        """Extract column lines data from the columns table"""
        column_lines = {
            'header': [],
            'items': [],
            'summary': []
        }

        # Iterate through all rows in the columns table
        for row in range(self.columns_table.rowCount()):
            # Get section, which should be in the first column
            section_item = self.columns_table.item(row, 0)
            if not section_item:
                continue

            section = section_item.text().lower()

            # Make sure section is valid
            if section not in column_lines:
                continue

            # Get coordinates
            try:
                table_idx_item = self.columns_table.item(row, 1)
                x_pos_item = self.columns_table.item(row, 2)
                desc_item = self.columns_table.item(row, 3)

                if table_idx_item and x_pos_item and desc_item:
                    table_idx = int(table_idx_item.text()) - 1  # Convert to 0-based index
                    x_pos = float(x_pos_item.text())

                    # Parse coordinates from description
                    desc = desc_item.text()
                    start_coords = desc.split("End:")[0].strip("Start: ()").split(",")
                    end_coords = desc.split("End:")[1].strip(" ()").split(",")

                    if len(start_coords) == 2 and len(end_coords) == 2:
                        start_x = float(start_coords[0].strip())
                        start_y = float(start_coords[1].strip())
                        end_x = float(end_coords[0].strip())
                        end_y = float(end_coords[1].strip())

                        # Create start and end points as dictionaries
                        start_point = {'x': start_x, 'y': start_y}
                        end_point = {'x': end_x, 'y': end_y}

                        # Add to column lines with table index
                        column_lines[section].append([start_point, end_point, table_idx])
                    else:
                        print(f"Warning: Invalid coordinate format in description: {desc}")
                continue

            except Exception as e:
                print(f"Error processing row {row}: {str(e)}")
                continue

        # If no column lines were found, preserve the original ones based on template type
        if not any(column_lines.values()) and hasattr(self, 'template_data'):
            print("No column lines found in table, preserving original column lines data")

            # Check if this is a multi-page template
            if self.template_data.get("template_type") == "multi":
                # For multi-page templates, get page-specific column lines
                page_column_lines = self.template_data.get("page_column_lines", [])
                if hasattr(self, 'current_page') and self.current_page < len(page_column_lines):
                    print(f"Preserving multi-page column lines for page {self.current_page}")
                    return page_column_lines[self.current_page]

            # Otherwise fallback to dual coordinate column lines (for single-page templates)
            drawing_column_lines = self.template_data.get('drawing_column_lines')
            if drawing_column_lines:
                print("Preserving single-page dual coordinate column lines")
                return drawing_column_lines
            else:
                print("No column lines found - returning empty")
                return {'header': [], 'items': [], 'summary': []}

        return column_lines

    def get_config_data(self):
        """Extract configuration data from the dialog"""
        config = {}

        # If we have a template_data, preserve any existing config data not overwritten
        if hasattr(self, 'template_data') and 'config' in self.template_data:
            # Start with a copy of the existing config to preserve any custom fields
            config = self.template_data['config'].copy()

        # Add extraction parameters to config with proper section-specific structure
        # Get global parameters
        split_text = self.split_text.isChecked()
        strip_text = self.strip_text.text().replace('\\n', '\n')
        flavor = 'stream'  # This is fixed

        # Create extraction parameters with section-specific parameters
        extraction_params = {
            'header': {
                'row_tol': self.header_row_tol.value(),
                'flavor': flavor,
                'split_text': split_text,
                'strip_text': strip_text,
                'edge_tol': 0.5  # Add edge_tol parameter
            },
            'items': {
                'row_tol': self.items_row_tol.value(),
                'flavor': flavor,
                'split_text': split_text,
                'strip_text': strip_text,
                'edge_tol': 0.5  # Add edge_tol parameter
            },
            'summary': {
                'row_tol': self.summary_row_tol.value(),
                'flavor': flavor,
                'split_text': split_text,
                'strip_text': strip_text,
                'edge_tol': 0.5  # Add edge_tol parameter
            },
            # Keep global parameters for backward compatibility
            'split_text': split_text,
            'strip_text': strip_text,
            'flavor': flavor
        }

        # Add multi-table mode
        config['multi_table_mode'] = self.multi_table_mode.isChecked()

        # Multi-page options removed - using simplified page-wise approach

        # Add extraction parameters
        config['extraction_params'] = extraction_params

        # Preserve original coordinates if they exist
        if hasattr(self, 'template_data') and 'config' in self.template_data:
            original_config = self.template_data['config']

            # Preserve original_regions if they exist
            if 'original_regions' in original_config:
                config['original_regions'] = original_config['original_regions']

            # Preserve original_column_lines if they exist
            if 'original_column_lines' in original_config:
                config['original_column_lines'] = original_config['original_column_lines']

            # Preserve scale_factors if they exist
            if 'scale_factors' in original_config:
                config['scale_factors'] = original_config['scale_factors']

            # Preserve store_original_coords flag if it exists
            if 'store_original_coords' in original_config:
                config['store_original_coords'] = original_config['store_original_coords']

            # Preserve any additional custom parameters
            known_params = ['multi_table_mode', 'extraction_params', 'regex_patterns', 'use_middle_page',
                           'fixed_page_count', 'total_pages', 'page_indices', 'store_original_coords',
                           'original_regions', 'original_column_lines', 'scale_factors']

            for key, value in original_config.items():
                if key not in known_params and key not in config:
                    config[key] = value

        # Preserve regex_patterns if they exist in the original config
        if hasattr(self, 'template_data') and 'config' in self.template_data:
            original_config = self.template_data['config']
            if 'regex_patterns' in original_config:
                config['regex_patterns'] = original_config['regex_patterns']
            else:
                # Add empty regex patterns structure for backward compatibility
                config['regex_patterns'] = {
                    'header': {},
                    'items': {},
                    'summary': {}
                }
        else:
            # Add empty regex patterns structure for backward compatibility
            config['regex_patterns'] = {
                'header': {},
                'items': {},
                'summary': {}
            }
        # For multi-page templates with page-specific config
        if (self.template_data.get("template_type") == "multi" and
            hasattr(self, 'page_specific_config') and
            self.page_specific_config.isChecked()):
            # Create or update page_configs
            page_configs = self.template_data.get('page_configs', [None] * self.page_count)

            # Ensure page_configs list is long enough
            while len(page_configs) < self.page_count:
                page_configs.append(None)

            # Create page-specific configs for each page
            for page_idx in range(self.page_count):
                # If this is the current page, use the current values
                if page_idx == self.current_page:
                    page_config = {
                        'extraction_params': extraction_params.copy()
                    }

                    # Add regex_patterns if they exist in the original config
                    if 'regex_patterns' in config:
                        page_config['regex_patterns'] = config['regex_patterns'].copy()
                    page_configs[page_idx] = page_config
                # Otherwise, keep existing config or initialize
                elif page_idx >= len(page_configs) or page_configs[page_idx] is None:
                    # Initialize with global config
                    page_config = {
                        'extraction_params': config['extraction_params'].copy()
                    }

                    # Add regex_patterns if they exist in the original config
                    if 'regex_patterns' in config:
                        page_config['regex_patterns'] = config['regex_patterns'].copy()

                    page_configs[page_idx] = page_config

            # Save page_configs to template_data
            config['page_configs'] = page_configs

        return config

    def get_mapping_config_data(self):
        """Extract mapping configuration data from the dialog"""
        mapping_config = {
            'approach': 'page_wise',  # Default
            'page_wise': {
                'first_page': 1,
                'middle_pages': 'sequential',
                'last_page': 'last_template_page'
            },
            'region_wise': {
                'header': {'source_page': '1'},
                'items': {'source_page': '1-n'},
                'summary': {'source_page': 'n'}
            }
        }

        try:
            # Get the current approach
            if hasattr(self, 'mapping_approach_combo'):
                approach_text = self.mapping_approach_combo.currentText()
                if approach_text == "Region-wise Mapping":
                    mapping_config['approach'] = 'region_wise'
                else:
                    mapping_config['approach'] = 'page_wise'

            # Get page-wise configuration
            if hasattr(self, 'first_page_spin'):
                mapping_config['page_wise']['first_page'] = self.first_page_spin.value()

            if hasattr(self, 'middle_pages_combo'):
                middle_text = self.middle_pages_combo.currentText()
                if middle_text == "Repeat Last":
                    mapping_config['page_wise']['middle_pages'] = 'repeat_last'
                elif middle_text == "Repeat First":
                    mapping_config['page_wise']['middle_pages'] = 'repeat_first'
                else:
                    mapping_config['page_wise']['middle_pages'] = 'sequential'

            if hasattr(self, 'last_page_combo'):
                last_text = self.last_page_combo.currentText()
                if last_text == "Same as First":
                    mapping_config['page_wise']['last_page'] = 'same_as_first'
                elif last_text == "Specific Page" and hasattr(self, 'last_page_spin'):
                    mapping_config['page_wise']['last_page'] = self.last_page_spin.value()
                else:
                    mapping_config['page_wise']['last_page'] = 'last_template_page'

            # Get region-wise configuration from text inputs
            if hasattr(self, 'header_source_input'):
                header_text = self.header_source_input.text().strip()
                mapping_config['region_wise']['header']['source_page'] = header_text or '1'

            if hasattr(self, 'items_source_input'):
                items_text = self.items_source_input.text().strip()
                mapping_config['region_wise']['items']['source_page'] = items_text or '1-n'

            if hasattr(self, 'summary_source_input'):
                summary_text = self.summary_source_input.text().strip()
                mapping_config['region_wise']['summary']['source_page'] = summary_text or 'n'

        except Exception as e:
            print(f"Error collecting mapping configuration: {e}")
            # Return default configuration on error

        return mapping_config

    # Removed duplicate validate_template_data function - using the one in TemplateManager class

    # Removed clone_regions_to_another_page function as it is no longer needed

    def display_extraction_parameters(self):
        """Display extraction parameters in the text edit"""
        try:
            extraction_text = "Extraction Parameters:\n"

            # Check for multi-page template with page-specific extraction parameters
            if (self.template_data.get('template_type') == 'multi' and
                'page_configs' in self.template_data and
                self.current_page < len(self.template_data['page_configs']) and
                self.template_data['page_configs'][self.current_page] and
                'extraction_params' in self.template_data['page_configs'][self.current_page]):

                # Get page-specific extraction parameters
                extraction_params = self.template_data['page_configs'][self.current_page]['extraction_params']
                extraction_text += f"\nPage {self.current_page + 1} Specific Parameters:\n"

                for section, params in extraction_params.items():
                    if isinstance(params, dict):
                        extraction_text += f"\n{section.capitalize()}:\n"
                        for param, value in params.items():
                            extraction_text += f"  {param}: {value}\n"

                # Add global parameters from page-specific config
                global_params = {k: v for k, v in extraction_params.items()
                               if not isinstance(v, dict)}
                if global_params:
                    extraction_text += "\nPage Global Parameters:\n"
                    for param, value in global_params.items():
                        extraction_text += f"  {param}: {value}\n"

                # Also show global extraction parameters if available
                if 'config' in self.template_data and 'extraction_params' in self.template_data['config']:
                    global_extraction_params = self.template_data['config']['extraction_params']
                    extraction_text += "\nTemplate Global Parameters:\n"

                    for section, params in global_extraction_params.items():
                        if isinstance(params, dict) and section not in extraction_params:
                            extraction_text += f"\n{section.capitalize()}:\n"
                            for param, value in params.items():
                                extraction_text += f"  {param}: {value}\n"

                    # Add global parameters that aren't in page-specific config
                    global_global_params = {k: v for k, v in global_extraction_params.items()
                                         if not isinstance(v, dict) and k not in global_params}
                    if global_global_params:
                        extraction_text += "\nTemplate Global Parameters:\n"
                        for param, value in global_global_params.items():
                            extraction_text += f"  {param}: {value}\n"

            # For single-page templates or if no page-specific parameters
            elif 'config' in self.template_data and 'extraction_params' in self.template_data['config']:
                extraction_params = self.template_data['config']['extraction_params']

                # Multi-page options removed - using simplified page-wise approach

                for section, params in extraction_params.items():
                    if isinstance(params, dict):
                        extraction_text += f"\n{section.capitalize()}:\n"
                        for param, value in params.items():
                            extraction_text += f"  {param}: {value}\n"

                # Add global parameters
                global_params = {k: v for k, v in extraction_params.items()
                               if not isinstance(v, dict)}
                if global_params:
                    extraction_text += "\nGlobal Parameters:\n"
                    for param, value in global_params.items():
                        extraction_text += f"  {param}: {value}\n"
            else:
                extraction_text += "\nNo extraction parameters found."

            self.extraction_params_text.setText(extraction_text)
        except Exception as e:
            print(f"Error displaying extraction parameters: {str(e)}")
            import traceback
            traceback.print_exc()
            self.extraction_params_text.setText(f"Error displaying extraction parameters: {str(e)}")

    def show_raw_config(self, config):
        """Show the raw configuration in a dialog for debugging"""
        try:
            import json
            config_text = json.dumps(config, indent=2)

            dialog = QDialog(self)
            dialog.setWindowTitle("Raw Configuration")
            dialog.setMinimumWidth(600)
            dialog.setMinimumHeight(400)

            layout = QVBoxLayout(dialog)

            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            text_edit.setFont(QFont("Courier New", 10))
            text_edit.setText(config_text)

            layout.addWidget(text_edit)

            close_btn = QPushButton("Close")
            close_btn.clicked.connect(dialog.close)
            layout.addWidget(close_btn)

            dialog.exec()
        except Exception as e:
            print(f"Error showing raw config: {e}")
            import traceback
            traceback.print_exc()

    # Removed clone_column_lines_to_another_page function as it is no longer needed

class TemplateManager(QWidget):
    """Widget for managing invoice templates"""

    template_selected = Signal(dict)  # Emits when a template is selected for use
    go_back = Signal()  # Signal to go back to previous screen

    def __init__(self, pdf_processor=None):
        # Ensure QApplication exists before creating widgets
        if QApplication.instance() is None:
            print("Creating QApplication instance because none exists")
            self.app = QApplication([])
        else:
            self.app = QApplication.instance()

        super().__init__()
        self.pdf_processor = pdf_processor
        self.db = InvoiceDatabase()

        # Initialize regions attribute to avoid 'can't set attribute' error
        self.regions = {
            'header': [],
            'items': [],
            'summary': []
        }

        # Set global stylesheet to ensure all text is visible
        self.setStyleSheet("""
            QWidget {
                color: black;
                background-color: white;
            }
            QLabel {
                color: #333333;
            }
            QTableWidgetItem {
                color: black;
            }
            QMessageBox {
                color: black;
            }
            QMessageBox QLabel {
                color: black;
            }
        """)

        self.initUI()
        self.load_templates()

    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # Title
        title = QLabel("Invoice Template Management")
        title.setFont(QFont("Arial", 24, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #333333; margin: 10px 0 5px 0;")
        layout.addWidget(title)

        # Navigation buttons below title
        nav_layout = QHBoxLayout()
        nav_layout.setContentsMargins(0, 0, 0, 5)

        back_btn = QPushButton("← Back")
        back_btn.clicked.connect(self.go_back.emit)
        back_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                color: #000000;
                padding: 5px 15px;
                border-radius: 4px;
                height: 25px;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
            }
        """)
        nav_layout.addWidget(back_btn)

        # Refresh List button next to Back button
        refresh_btn = QPushButton("Refresh List")
        refresh_btn.clicked.connect(self.load_templates)
        refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px 15px;
                border-radius: 4px;
                height: 25px;
            }
            QPushButton:hover {
                background-color: #3E8E41;
            }
        """)
        nav_layout.addWidget(refresh_btn)

        nav_layout.addStretch()

        # Add Template button
        add_btn = QPushButton("Add Template")
        add_btn.clicked.connect(self.add_template)
        add_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 5px 15px;
                border-radius: 4px;
                height: 25px;
            }
            QPushButton:hover {
                background-color: #3158D3;
            }
        """)
        nav_layout.addWidget(add_btn)

        # Export Template button
        export_btn = QPushButton("Export Template")
        export_btn.clicked.connect(self.export_template)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 5px 15px;
                border-radius: 4px;
                height: 25px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        nav_layout.addWidget(export_btn)

        # Import Template button
        import_btn = QPushButton("Import Template")
        import_btn.clicked.connect(self.import_template)
        import_btn.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
                color: white;
                padding: 5px 15px;
                border-radius: 4px;
                height: 25px;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
        """)
        nav_layout.addWidget(import_btn)

        # Add Reset Database button
        reset_db_btn = QPushButton("Reset Database")
        reset_db_btn.clicked.connect(self.reset_database)
        reset_db_btn.setStyleSheet("""
            QPushButton {
                background-color: #ffaaaa;
                color: #aa0000;
                padding: 5px 15px;
                border-radius: 4px;
                font-weight: bold;
                height: 25px;
            }
            QPushButton:hover {
                background-color: #ff8888;
            }
        """)
        nav_layout.addWidget(reset_db_btn)

        layout.addLayout(nav_layout)

        # Description
        description = QLabel("Create, manage and apply invoice extraction templates")
        description.setWordWrap(True)
        description.setStyleSheet("color: #666666; font-size: 16px; margin: 5px 0;")
        description.setAlignment(Qt.AlignCenter)
        layout.addWidget(description)

        # No actions toolbar needed anymore as all buttons are in the navigation bar

        # Templates table
        table_container = QFrame()
        table_container.setFrameShape(QFrame.StyledPanel)
        table_container.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)

        table_layout = QVBoxLayout(table_container)
        table_layout.setContentsMargins(10, 10, 10, 10)

        # Add a header label for the table
        templates_header = QLabel("Your Templates")
        templates_header.setFont(QFont("Arial", 14, QFont.Bold))
        templates_header.setStyleSheet("color: #333; margin-bottom: 10px;")
        table_layout.addWidget(templates_header)

        self.templates_table = QTableWidget()
        self.templates_table.setColumnCount(5)
        self.templates_table.setHorizontalHeaderLabels(["Name", "Description", "Type", "Created Date/Time", "Actions"])
        self.templates_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.templates_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.templates_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.templates_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)  # Allow user to resize the date/time column
        self.templates_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)

        # Set a minimum width for the date/time column to ensure it can display the full date and time
        self.templates_table.setColumnWidth(3, 200)
        self.templates_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.templates_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.templates_table.setAlternatingRowColors(True)
        self.templates_table.verticalHeader().setVisible(False)
        self.templates_table.setShowGrid(True)
        self.templates_table.setStyleSheet("""
            QTableWidget {
                border: none;
                gridline-color: #e0e0e0;
                selection-background-color: #f0f7ff;
                selection-color: #000;
                color: #000000; /* Ensuring text is black */
                background-color: white;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
                color: #000000; /* Ensuring text is black */
                background-color: white;
            }
            QTableWidget::item:selected {
                background-color: #f0f7ff;
                color: #000000; /* Ensuring selected text is black */
            }
            QHeaderView::section {
                background-color: #f8f8f8;
                padding: 8px;
                border: none;
                border-bottom: 2px solid #e0e0e0;
                font-weight: bold;
                color: #333;
                font-size: 13px;
            }
            QTableWidget::item:alternate {
                background-color: #f9f9f9;
                color: #000000; /* Ensuring text is black */
            }
            /* Fix for action buttons in table cells */
            QTableWidget QWidget {
                background-color: transparent;
            }
            QTableWidget::item:alternate:selected {
                background-color: #f0f7ff;
                color: #000000; /* Ensuring selected text is black */
            }
        """)

        table_layout.addWidget(self.templates_table)

        layout.addWidget(table_container)

        # Navigation buttons are now at the top

        self.setLayout(layout)

    def get_template_id_from_row(self, row):
        """Get the template ID for the given row"""
        if 0 <= row < self.templates_table.rowCount():
            # Assuming template ID is stored in the table or can be retrieved from the database
            templates = self.db.get_all_templates()
            if row < len(templates):
                return templates[row]["id"]
        return None

    def show_context_menu(self, position):
        """Show context menu for the templates table"""
        menu = QMenu(self)

        edit_action = menu.addAction("Edit")
        delete_action = menu.addAction("Delete")

        # Get the row under the cursor
        row = self.templates_table.rowAt(position.y())

        # Only enable actions if a valid row is clicked
        edit_action.setEnabled(row >= 0)
        delete_action.setEnabled(row >= 0)

        # Connect actions to slots with lambda functions that ignore the 'checked' parameter
        if row >= 0:
            template_id = self.get_template_id_from_row(row)
            if template_id:
                edit_action.triggered.connect(lambda checked=False: self.edit_template(template_id))
                delete_action.triggered.connect(lambda checked=False: self.delete_template(template_id))

        menu.exec_(self.templates_table.mapToGlobal(position))

    def edit_template(self, template_id):
        """Edit the selected template settings and configuration"""
        try:
            # Get template for editing
            template = self.db.get_template(template_id=template_id)
            if not template:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Template Not Found")
                msg_box.setText("The selected template could not be found.")
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setStyleSheet("QLabel { color: black; }")
                msg_box.exec()
                return

            # Show edit dialog
            dialog = EditTemplateDialog(self, template_data=template)

            # Set a more descriptive window title
            dialog.setWindowTitle(f"Edit Template - {template['name']}")

            if dialog.exec() == QDialog.Accepted:
                try:
                    # Get updated data with proper error handling
                    updated_data = dialog.get_template_data()

                    # Validate that we have all required data
                    if not self.validate_template_data(updated_data):
                        print("Template data validation failed, aborting update")
                        return

                    new_name = updated_data["name"]
                    new_description = updated_data["description"]

                    if not new_name:
                        msg_box = QMessageBox(self)
                        msg_box.setWindowTitle("Invalid Name")
                        msg_box.setText("Please provide a valid template name.")
                        msg_box.setIcon(QMessageBox.Warning)
                        msg_box.setStyleSheet("QLabel { color: black; }")
                        msg_box.exec()
                        return

                    # Create a progress dialog instead of a message box
                    progress = QProgressDialog("Updating template...", None, 0, 100, self)
                    progress.setWindowTitle("Updating Template")
                    progress.setWindowModality(Qt.WindowModal)
                    progress.setMinimumDuration(0)  # Show immediately
                    progress.setValue(0)
                    progress.setAutoClose(True)
                    progress.setAutoReset(True)
                    progress.setCancelButton(None)  # No cancel button
                    progress.setFixedSize(300, 100)
                    progress.setStyleSheet("QLabel { color: black; }")
                    progress.show()

                    # Process events to ensure the dialog is displayed
                    QApplication.processEvents()

                    try:
                        # Update progress to show we've started
                        progress.setValue(10)
                        QApplication.processEvents()

                        # Debug output
                        print(f"\nAttempting to save template: {new_name}")
                        print(f"Template regions: {len(updated_data['regions'].get('header', []))} header, {len(updated_data['regions'].get('items', []))} items, {len(updated_data['regions'].get('summary', []))} summary")
                        print(f"Config has regex_patterns: {'regex_patterns' in updated_data['config']}")

                        # Update progress
                        progress.setValue(30)
                        QApplication.processEvents()

                        # Check if new name exists (if changed)
                        if new_name != template["name"]:
                            # Need to delete old template and create new one with new name
                            print(f"Name changed from '{template['name']}' to '{new_name}', deleting old template")
                            self.db.delete_template(template_id=template_id)

                            # Update progress
                            progress.setValue(50)
                            QApplication.processEvents()

                            # Save the new template with appropriate data based on template type
                            if updated_data["template_type"] == "multi":
                                # For multi-page templates, include page-specific data
                                print("\nCreating new multi-page template with the following data:")
                                print(f"- Page count: {updated_data['page_count']}")

                                # Log regions data
                                page_regions = updated_data.get("page_regions", [])
                                print(f"- Page regions: {len(page_regions)} pages")
                                for i, page_region in enumerate(page_regions):
                                    region_counts = {section: len(rects) for section, rects in page_region.items()}
                                    print(f"  - Page {i+1}: {region_counts}")

                                # Log column lines data
                                page_column_lines = updated_data.get("page_column_lines", [])
                                print(f"- Page column lines: {len(page_column_lines)} pages")
                                for i, page_column_line in enumerate(page_column_lines):
                                    column_counts = {section: len(lines) for section, lines in page_column_line.items()}
                                    print(f"  - Page {i+1}: {column_counts}")

                                # Create template with page-specific data
                                new_id = self.db.save_template(
                                    name=new_name,
                                    description=new_description,
                                    config=updated_data["config"],
                                    template_type=updated_data["template_type"],
                                    page_count=updated_data["page_count"],
                                    json_template=updated_data.get("json_template"),
                                    drawing_regions=updated_data.get("drawing_regions"),
                                    drawing_column_lines=updated_data.get("drawing_column_lines"),
                                    extraction_regions=updated_data.get("extraction_regions"),
                                    extraction_column_lines=updated_data.get("extraction_column_lines"),
                                    drawing_page_regions=updated_data.get("drawing_page_regions"),
                                    drawing_page_column_lines=updated_data.get("drawing_page_column_lines"),
                                    extraction_page_regions=updated_data.get("extraction_page_regions"),
                                    extraction_page_column_lines=updated_data.get("extraction_page_column_lines")
                                )
                            else:
                                # For single-page templates, use the standard fields
                                print("\nCreating new single-page template with the following data:")
                                region_counts = {section: len(rects) for section, rects in updated_data["regions"].items()}
                                print(f"- Regions: {region_counts}")
                                column_counts = {section: len(lines) for section, lines in updated_data["column_lines"].items()}
                                print(f"- Column lines: {column_counts}")

                            new_id = self.db.save_template(
                                name=new_name,
                                description=new_description,
                                config=updated_data["config"],
                                template_type=updated_data["template_type"],
                                json_template=updated_data.get("json_template"),
                                drawing_regions=updated_data.get("drawing_regions"),
                                drawing_column_lines=updated_data.get("drawing_column_lines"),
                                extraction_regions=updated_data.get("extraction_regions"),
                                extraction_column_lines=updated_data.get("extraction_column_lines")
                            )
                            print(f"Created new template with ID: {new_id}")
                        else:
                            # Just update the template data
                            print(f"Updating existing template with ID: {template_id}")

                            # Update progress
                            progress.setValue(50)
                            QApplication.processEvents()

                            # Update the template with dual coordinate data
                            if updated_data["template_type"] == "multi":
                                # For multi-page templates, include page-specific dual coordinate data
                                self.db.save_template(
                                    name=new_name,
                                    description=new_description,
                                    regions={'header': [], 'items': [], 'summary': []},  # Legacy - empty
                                    column_lines={'header': [], 'items': [], 'summary': []},  # Legacy - empty
                                    config=updated_data["config"],
                                    template_type=updated_data["template_type"],
                                    page_count=updated_data["page_count"],
                                    page_regions=updated_data.get("page_regions", []),
                                    page_column_lines=updated_data.get("page_column_lines", []),
                                    page_configs=updated_data["config"].get("page_configs", []),
                                    json_template=updated_data.get("json_template"),
                                    drawing_page_regions=updated_data.get("drawing_page_regions"),
                                    drawing_page_column_lines=updated_data.get("drawing_page_column_lines"),
                                    extraction_page_regions=updated_data.get("extraction_page_regions"),
                                    extraction_page_column_lines=updated_data.get("extraction_page_column_lines")
                                )
                            else:
                                # For single-page templates, use dual coordinate data
                                self.db.save_template(
                                    name=new_name,
                                    description=new_description,
                                    config=updated_data["config"],
                                    template_type=updated_data["template_type"],
                                    json_template=updated_data.get("json_template"),
                                    drawing_regions=updated_data.get("drawing_regions"),
                                    drawing_column_lines=updated_data.get("drawing_column_lines"),
                                    extraction_regions=updated_data.get("extraction_regions"),
                                    extraction_column_lines=updated_data.get("extraction_column_lines")
                                )

                        # Update progress to completion
                        progress.setValue(100)
                        QApplication.processEvents()

                        # Make sure dialog is closed
                        progress.close()
                        QApplication.processEvents()

                        # Show success message with details
                        success_message = f"""
<h3>Template Updated Successfully</h3>
<p>The template has been updated with the following information:</p>
<ul>
    <li><b>Name:</b> {new_name}</li>
    <li><b>Description:</b> {new_description or "No description"}</li>
    <li><b>Type:</b> {updated_data["template_type"].title()}</li>
    <li><b>Multi-table Mode:</b> {"Enabled" if updated_data["config"]["multi_table_mode"] else "Disabled"}</li>
    <li><b>JSON Template:</b> {"<span style='color: green;'>✓ Included</span>" if updated_data.get("json_template") else "<span style='color: gray;'>Not included</span>"}</li>
</ul>
"""

                        # Add regions information based on template type
                        success_message += "<p><b>Regions:</b></p><ul>"
                        if updated_data["template_type"] == "multi":
                            # For multi-page templates, show page-specific regions
                            page_regions = updated_data.get("page_regions", [])
                            for page_idx, page_region in enumerate(page_regions):
                                success_message += f"<li><b>Page {page_idx + 1}:</b><ul>"
                                for section, rects in page_region.items():
                                    success_message += f"<li>{section.title()}: {len(rects)} table(s)</li>"
                                success_message += "</ul></li>"
                        else:
                            # For single-page templates, use the regular regions
                            for section, rects in updated_data["regions"].items():
                                success_message += f"<li>{section.title()}: {len(rects)} table(s)</li>"
                        success_message += "</ul>"

                        # Add column lines information based on template type
                        success_message += "<p><b>Column Lines:</b></p><ul>"
                        if updated_data["template_type"] == "multi":
                            # For multi-page templates, show page-specific column lines
                            page_column_lines = updated_data.get("page_column_lines", [])
                            for page_idx, page_column_line in enumerate(page_column_lines):
                                success_message += f"<li><b>Page {page_idx + 1}:</b><ul>"
                                for section, lines in page_column_line.items():
                                    success_message += f"<li>{section.title()}: {len(lines)} line(s)</li>"
                                success_message += "</ul></li>"
                        else:
                            # For single-page templates, use the regular column lines
                            for section, lines in updated_data["column_lines"].items():
                                success_message += f"<li>{section.title()}: {len(lines)} line(s)</li>"
                            success_message += "</ul>"

                        # Add regex pattern information if available
                        if 'regex_patterns' in updated_data['config']:
                            success_message += "<p><b>Regex Patterns:</b></p><ul>"
                            for section, patterns in updated_data['config']['regex_patterns'].items():
                                pattern_list = []
                                for pattern_type, pattern in patterns.items():
                                    if pattern:
                                        pattern_list.append(f"{pattern_type}: '{pattern}'")
                                if pattern_list:
                                    success_message += f"<li>{section.title()}: {', '.join(pattern_list)}</li>"
                            success_message += "</ul>"

                        success_msg = QMessageBox(self)
                        success_msg.setWindowTitle("Template Updated")
                        success_msg.setText("Template Updated")
                        success_msg.setInformativeText(success_message)
                        success_msg.setIcon(QMessageBox.Information)
                        success_msg.setStyleSheet("QLabel { color: black; }")
                        success_msg.exec()

                        # Refresh the template list
                        self.load_templates()

                    except sqlite3.Error as db_e:
                        # Handle database-specific errors
                        progress.close()
                        QApplication.processEvents()

                        error_message = f"""
<h3>Database Error</h3>
<p>A database error occurred while trying to save the template:</p>
<p style='color: #D32F2F;'>{str(db_e)}</p>
<p>This might be due to database corruption, permissions issues, or disk space limitations.</p>
"""
                        error_dialog = QMessageBox(self)
                        error_dialog.setWindowTitle("Database Error")
                        error_dialog.setText("Error Saving Template")
                        error_dialog.setInformativeText(error_message)
                        error_dialog.setIcon(QMessageBox.Critical)
                        error_dialog.setStyleSheet("QLabel { color: black; }")
                        error_dialog.exec()

                        print(f"Database error in edit_template: {str(db_e)}")
                        import traceback
                        traceback.print_exc()

                    except Exception as inner_e:
                        # Close the progress dialog for any other error
                        progress.close()
                        QApplication.processEvents()

                        error_message = f"""
<h3>Error Saving Template</h3>
<p>An error occurred while saving the template:</p>
<p style='color: #D32F2F;'>{str(inner_e)}</p>
<p>The template may not have been updated properly.</p>
"""
                        error_dialog = QMessageBox(self)
                        error_dialog.setWindowTitle("Save Error")
                        error_dialog.setText("Error Saving Template")
                        error_dialog.setInformativeText(error_message)
                        error_dialog.setIcon(QMessageBox.Critical)
                        error_dialog.setStyleSheet("QLabel { color: black; }")
                        error_dialog.exec()

                        print(f"Error in edit_template save operation: {str(inner_e)}")
                        import traceback
                        traceback.print_exc()

                    finally:
                        # Make sure progress dialog is closed in all cases
                        try:
                            progress.close()
                            QApplication.processEvents()
                        except Exception as close_e:
                            print(f"Error closing progress dialog: {str(close_e)}")

                except AttributeError as attr_e:
                    error_message = f"""
<h3>Template Update Error</h3>
<p>There was a problem accessing template data:</p>
<p style='color: #D32F2F;'>{str(attr_e)}</p>
<p>This could be due to missing or corrupted template information.</p>
"""
                    error_dialog = QMessageBox(self)
                    error_dialog.setWindowTitle("Error")
                    error_dialog.setText("Error Updating Template")
                    error_dialog.setInformativeText(error_message)
                    error_dialog.setIcon(QMessageBox.Critical)
                    error_dialog.setStyleSheet("QLabel { color: black; }")
                    error_dialog.exec()

                    print(f"AttributeError in edit_template: {str(attr_e)}")
                    import traceback
                    traceback.print_exc()

                except ValueError as val_e:
                    error_message = f"""
<h3>Template Value Error</h3>
<p>There was a problem with the template data values:</p>
<p style='color: #D32F2F;'>{str(val_e)}</p>
<p>Please check that all fields contain valid information.</p>
"""
                    error_dialog = QMessageBox(self)
                    error_dialog.setWindowTitle("Error")
                    error_dialog.setText("Error Updating Template")
                    error_dialog.setInformativeText(error_message)
                    error_dialog.setIcon(QMessageBox.Critical)
                    error_dialog.setStyleSheet("QLabel { color: black; }")
                    error_dialog.exec()

                    print(f"ValueError in edit_template: {str(val_e)}")
                    import traceback
                    traceback.print_exc()

        except Exception as e:
            error_message = f"""
<h3>Error Updating Template</h3>
<p>An error occurred while trying to update the template:</p>
<p style='color: #D32F2F;'>{str(e)}</p>
<p>Please try again or contact support if this issue persists.</p>
"""
            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Error Updating Template")
            error_dialog.setInformativeText(error_message)
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setStyleSheet("QLabel { color: black; }")
            error_dialog.exec()

            # Print detailed error information to help with debugging
            print(f"Error in edit_template: {str(e)}")
            import traceback
            traceback.print_exc()

    def delete_template(self, template_id):
        """Delete the selected template from the database"""
        try:
            # Get template name for confirmation
            template = self.db.get_template(template_id=template_id)
            if not template:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Template Not Found")
                msg_box.setText("The selected template could not be found.")
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setStyleSheet("QLabel { color: black; }")
                msg_box.exec()
                return

            # Create a more detailed confirmation message
            confirmation_message = f"""
<h3>Confirm Template Deletion</h3>

<p style='color: #D32F2F;'><b>Warning:</b> This action cannot be undone.</p>

<p>You are about to delete the following template:</p>
<ul>
    <li><b>Name:</b> {template.get('name', 'Unnamed Template')}</li>
    <li><b>Type:</b> {template.get('template_type', 'Unknown').title()}</li>
    <li><b>Regions:</b> {sum(len(rects) for rects in template.get('regions', {}).values())} table(s)</li>
</ul>

<p>Are you sure you want to proceed?</p>
"""

            # Create a custom confirmation dialog
            confirm_dialog = QMessageBox(self)
            confirm_dialog.setWindowTitle("Confirm Deletion")
            confirm_dialog.setText("Delete Template?")
            confirm_dialog.setInformativeText(confirmation_message)
            confirm_dialog.setIcon(QMessageBox.Warning)
            confirm_dialog.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            confirm_dialog.setDefaultButton(QMessageBox.No)
            confirm_dialog.setStyleSheet("QLabel { color: black; }")

            # Set button texts
            yes_button = confirm_dialog.button(QMessageBox.Yes)
            yes_button.setText("Delete")
            no_button = confirm_dialog.button(QMessageBox.No)
            no_button.setText("Cancel")

            # Show the dialog
            result = confirm_dialog.exec()

            if result == QMessageBox.Yes:
                # Delete the template
                self.db.delete_template(template_id=template_id)

                # Show success message
                success_msg = QMessageBox(self)
                success_msg.setWindowTitle("Template Deleted")
                success_msg.setText("Template Successfully Deleted")
                success_msg.setInformativeText(f"The template '{template.get('name', 'Unnamed Template')}' has been permanently deleted.")
                success_msg.setIcon(QMessageBox.Information)
                success_msg.setStyleSheet("QLabel { color: black; }")
                success_msg.exec()

                # Refresh the template list
                self.load_templates()

        except Exception as e:
            error_message = f"""
<h3>Error Deleting Template</h3>
<p>An error occurred while trying to delete the template:</p>
<p style='color: #D32F2F;'>{str(e)}</p>
<p>Please try again or contact support if this issue persists.</p>
"""
            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Error Deleting Template")
            error_dialog.setInformativeText(error_message)
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setStyleSheet("QLabel { color: black; }")
            error_dialog.exec()

    def load_templates(self):
        """Load all templates from the database and display them in the table"""
        templates = self.db.get_all_templates()

        self.templates_table.setRowCount(len(templates))

        for row, template in enumerate(templates):
            # Template name
            name_item = QTableWidgetItem(template["name"])
            name_item.setToolTip(template["name"])
            name_item.setForeground(Qt.black)  # Explicitly set text color to black
            self.templates_table.setItem(row, 0, name_item)

            # Description
            desc_item = QTableWidgetItem(template["description"] if template["description"] else "")
            desc_item.setToolTip(template["description"] if template["description"] else "")
            desc_item.setForeground(Qt.black)  # Explicitly set text color to black
            self.templates_table.setItem(row, 1, desc_item)

            # Type with page count for multi-page templates
            if template["template_type"] == "multi":
                page_count = template.get("page_count", 1)
                type_text = f"Multi-page ({page_count} pages)"
            else:
                type_text = "Single-page"

            type_item = QTableWidgetItem(type_text)
            type_item.setForeground(Qt.black)  # Explicitly set text color to black
            self.templates_table.setItem(row, 2, type_item)

            # Created date - Format for better readability with 12-hour time format
            try:
                # Try to parse and format the date with time in 12-hour format
                date_str = template["creation_date"]
                if "T" in date_str:  # ISO format
                    # Split into date and time parts
                    date_time_parts = date_str.split("T")
                    date_parts = date_time_parts[0].split("-")

                    # Format date as DD/MM/YYYY
                    formatted_date = f"{date_parts[2]}/{date_parts[1]}/{date_parts[0]}"

                    # Add time in 12-hour format if available
                    if len(date_time_parts) > 1:
                        time_part = date_time_parts[1].split(".")[0]  # Remove milliseconds if present
                        time_parts = time_part.split(":")
                        if len(time_parts) >= 2:
                            # Convert to 12-hour format with AM/PM
                            hour = int(time_parts[0])
                            minute = time_parts[1]
                            am_pm = "AM" if hour < 12 else "PM"

                            # Convert hour from 24-hour to 12-hour format
                            if hour == 0:
                                hour = 12  # 00:00 becomes 12:00 AM
                            elif hour > 12:
                                hour = hour - 12  # 13:00 becomes 1:00 PM

                            formatted_time = f"{hour}:{minute} {am_pm}"
                            formatted_date = f"{formatted_date} {formatted_time}"
                else:
                    formatted_date = date_str
            except:
                formatted_date = template["creation_date"]

            date_item = QTableWidgetItem(formatted_date)
            date_item.setTextAlignment(Qt.AlignCenter)
            date_item.setForeground(Qt.black)  # Explicitly set text color to black
            self.templates_table.setItem(row, 3, date_item)

            # Action buttons
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(4, 0, 4, 0)
            actions_layout.setSpacing(8)
            actions_layout.setAlignment(Qt.AlignCenter)

            # Set explicit background color for the widget
            if row % 2 == 0:
                actions_widget.setStyleSheet("background-color: white;")
            else:
                actions_widget.setStyleSheet("background-color: #f9f9f9;")

            # Apply button with icon
            apply_btn = QPushButton("Apply")
            apply_btn.setProperty("template_id", template["id"])
            apply_btn.setProperty("action", "apply")
            apply_btn.setToolTip(f"Apply template: {template['name']}")
            apply_btn.clicked.connect(lambda checked=False, tid=template["id"]: self.apply_template(tid))
            apply_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4169E1;
                    color: white;
                    padding: 5px 12px;
                    border-radius: 4px;
                    font-weight: bold;
                    min-width: 70px;
                }
                QPushButton:hover {
                    background-color: #3158D3;
                }
            """)

            # Edit button with icon
            edit_btn = QPushButton("Edit")
            edit_btn.setProperty("template_id", template["id"])
            edit_btn.setProperty("action", "edit")
            edit_btn.setToolTip(f"Edit template: {template['name']}")
            edit_btn.clicked.connect(lambda checked=False, tid=template["id"]: self.edit_template(tid))
            edit_btn.setStyleSheet("""
                QPushButton {
                    background-color: #FF9800;
                    color: white;
                    padding: 5px 12px;
                    border-radius: 4px;
                    font-weight: bold;
                    min-width: 70px;
                }
                QPushButton:hover {
                    background-color: #F57C00;
                }
            """)

            # Delete button with icon
            delete_btn = QPushButton("Delete")
            delete_btn.setProperty("template_id", template["id"])
            delete_btn.setProperty("action", "delete")
            delete_btn.setToolTip(f"Delete template: {template['name']}")
            delete_btn.clicked.connect(lambda checked=False, tid=template["id"]: self.delete_template(tid))
            delete_btn.setStyleSheet("""
                QPushButton {
                    background-color: #D32F2F;
                    color: white;
                    padding: 5px 12px;
                    border-radius: 4px;
                    font-weight: bold;
                    min-width: 70px;
                }
                QPushButton:hover {
                    background-color: #B71C1C;
                }
            """)

            actions_layout.addWidget(apply_btn)
            actions_layout.addWidget(edit_btn)
            actions_layout.addWidget(delete_btn)

            # Set fixed height for the actions widget to ensure proper display
            actions_widget.setFixedHeight(40)

            # Set the cell widget
            self.templates_table.setCellWidget(row, 4, actions_widget)

        # If no templates, show a message
        if len(templates) == 0:
            self.templates_table.setRowCount(1)
            no_templates_item = QTableWidgetItem("No templates found. Create your first template!")
            no_templates_item.setTextAlignment(Qt.AlignCenter)
            no_templates_item.setForeground(Qt.black)  # Explicitly set text color to black
            self.templates_table.setItem(0, 0, no_templates_item)
            self.templates_table.setSpan(0, 0, 1, 5)

            # Set background color for the empty row
            for col in range(5):
                if col == 0:  # Skip the first column as it already has the message
                    continue
                empty_item = QTableWidgetItem("")
                empty_item.setBackground(QColor(240, 240, 240))  # Light gray background
                self.templates_table.setItem(0, col, empty_item)

        # Adjust row heights for better spacing
        for row in range(self.templates_table.rowCount()):
            self.templates_table.setRowHeight(row, 50)  # Increased height for better button display

    def apply_template(self, template_id):
        """Apply the selected template to the current PDF processor"""
        try:
            from PySide6.QtCore import QRect, QPoint

            # Get the template from the database - this will be the source of truth
            db_template = self.db.get_template(template_id=template_id)
            if not db_template:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Template Not Found")
                msg_box.setText("The selected template could not be found.")
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setStyleSheet("QLabel { color: black; }")
                msg_box.exec()
                return

            # Use the database template as our primary template
            template = db_template
            print(f"[DEBUG] Using template from database: {template['name']}")

            # Show a preview of the template settings
            template_preview = f"""
<h3>Template: {template['name']}</h3>

<p><b>Description:</b> {template['description'] or 'No description'}</p>
<p><b>Type:</b> {"Multi-page" if template.get('template_type') == 'multi' else "Single-page"}</p>
"""

            # Show additional parameters if they exist in the config
            if 'config' in template and isinstance(template['config'], dict):
                config = template['config']
                additional_params = []

                # Original coordinates are always stored by default, no need to show this parameter

                # Check for any custom parameters (exclude known parameters)
                known_params = ['multi_table_mode', 'extraction_params', 'regex_patterns', 'use_middle_page',
                               'fixed_page_count', 'total_pages', 'page_indices', 'store_original_coords',
                               'original_regions', 'original_column_lines', 'scale_factors']

                for key, value in config.items():
                    if key not in known_params:
                        # Format the value based on its type
                        if isinstance(value, bool):
                            formatted_value = 'Yes' if value else 'No'
                        elif isinstance(value, (dict, list)):
                            formatted_value = f"<i>(complex value)</i>"
                        else:
                            formatted_value = str(value)

                        additional_params.append(f"<li>{key}: {formatted_value}</li>")

                # Add additional parameters section if any were found
                if additional_params:
                    template_preview += "<p><b>Additional Parameters:</b></p><ul>"
                    template_preview += '\n'.join(additional_params)
                    template_preview += "</ul>"

            if template.get('template_type') == 'multi':
                template_preview += f"<p><b>Number of Pages:</b> {template.get('page_count', 1)}</p>"

                # For multi-page templates, show info for each page
                template_preview += "<p><b>Pages:</b></p><ul>"

                # Get page-specific data if available
                page_regions = template.get('page_regions', [])
                page_column_lines = template.get('page_column_lines', [])

                for page_idx in range(template.get('page_count', 1)):
                    template_preview += f"<li><b>Page {page_idx + 1}:</b><ul>"

                    # Regions for this page
                    if page_idx < len(page_regions):
                        regions = page_regions[page_idx]
                        region_count = sum(len(rects) for rects in regions.values())
                        template_preview += f"<li>Regions: {region_count}</li>"

                        # Add details about regions
                        if regions:
                            template_preview += "<ul>"
                            for section, rects in regions.items():
                                if rects:
                                    template_preview += f"<li>{section.title()}: {len(rects)} table(s)</li>"
                            template_preview += "</ul>"

                    # Column lines for this page
                    if page_idx < len(page_column_lines):
                        column_lines = page_column_lines[page_idx]
                        column_line_count = sum(len(lines) for lines in column_lines.values())
                        template_preview += f"<li>Column Lines: {column_line_count}</li>"

                    template_preview += "</ul></li>"

                template_preview += "</ul>"
            else:
                # Show information for single-page template
                region_count = 0
                if 'regions' in template and template['regions']:
                    template_preview += "<p><b>Tables:</b></p><ul>"
                    for section, rects in template['regions'].items():
                        if rects:
                            region_count += len(rects)
                            template_preview += f"<li>{section.title()}: {len(rects)} table(s)</li>"
                    template_preview += f"</ul><p><b>Total Regions:</b> {region_count}</p>"

                # Column lines info
                # Check dual coordinate column lines first, then legacy
                drawing_column_lines = template.get('drawing_column_lines', template.get('column_lines', {}))
                has_column_lines = any(drawing_column_lines.values()) if drawing_column_lines else False
                template_preview += f"<p><b>Column Lines:</b> {'Yes' if has_column_lines else 'No'}</p>"

            # Show configuration info
            template_preview += f"<p><b>Multi-table Mode:</b> {'Yes' if template.get('config', {}).get('multi_table_mode', False) else 'No'}</p>"

            # Show the preview dialog
            preview = QMessageBox(self)
            preview.setWindowTitle("Template Preview")
            preview.setText("Template Preview")
            preview.setInformativeText(template_preview)
            preview.setStandardButtons(QMessageBox.Apply | QMessageBox.Cancel)
            preview.setDefaultButton(QMessageBox.Apply)
            preview.setStyleSheet("QLabel { color: black; }")

            # If the user confirms, apply the template
            if preview.exec() == QMessageBox.Apply:
                # Clear the existing PDF in pdf_processor if it exists
                if self.pdf_processor and hasattr(self.pdf_processor, 'clear_all'):
                    print("Clearing existing PDF and regions before applying template")
                    self.pdf_processor.clear_all()

                # Helper function to convert coordinates from database format to PDF format
                def convert_coordinates(rect_data, pdf_height=None):
                    """
                    Convert coordinates from database format to PDF format (QRect)
                    Handles both original unscaled (x,y,width,height) and scaled (x1,y1,x2,y2) formats
                    Prioritizes using the original unscaled coordinates if available
                    """
                    if isinstance(rect_data, dict):
                        # First check if we have original unscaled coordinates
                        if 'x' in rect_data and 'y' in rect_data and 'width' in rect_data and 'height' in rect_data:
                            # Use the original unscaled coordinates (x,y,width,height) directly without conversion
                            x = float(rect_data['x'])
                            y = float(rect_data['y'])
                            width = float(rect_data['width'])
                            height = float(rect_data['height'])

                            # Ensure width and height are positive
                            if width < 0:
                                x += width
                                width = abs(width)
                            if height < 0:
                                y += height
                                height = abs(height)

                            print(f"Using original unscaled coordinates directly: x={x}, y={y}, width={width}, height={height}")
                            return QRect(int(x), int(y), int(width), int(height))

                        # Handle scaled format only (x1,y1,x2,y2)
                        elif 'x1' in rect_data and 'y1' in rect_data and 'x2' in rect_data and 'y2' in rect_data:
                            # Get coordinates from scaled database format
                            x1 = float(rect_data['x1'])
                            y1 = float(rect_data['y1'])
                            x2 = float(rect_data['x2'])
                            y2 = float(rect_data['y2'])

                            # Calculate width and height
                            width = x2 - x1
                            height = y2 - y1

                            # Ensure width and height are positive
                            if width < 0:
                                x1 += width
                                width = abs(width)
                            if height < 0:
                                y1 += height
                                height = abs(height)

                            print(f"Using scaled coordinates: x1={x1}, y1={y1}, width={width}, height={height}")
                            return QRect(int(x1), int(y1), int(width), int(height))
                        else:
                            print(f"Warning: Unknown rect format in template: {rect_data}")
                            return QRect()
                    elif isinstance(rect_data, QRect):
                        # Already a QRect, ensure width and height are positive
                        x = rect_data.x()
                        y = rect_data.y()
                        width = rect_data.width()
                        height = rect_data.height()

                        # Ensure width and height are positive
                        if width < 0:
                            x += width
                            width = abs(width)
                        if height < 0:
                            y += height
                            height = abs(height)

                        return QRect(x, y, width, height)
                    elif hasattr(rect_data, 'rect') and hasattr(rect_data, 'label'):
                        # StandardRegion object - extract the UI coordinates
                        from standardized_coordinates import StandardRegion
                        if isinstance(rect_data, StandardRegion):
                            ui_rect = rect_data.rect
                            x = ui_rect.x()
                            y = ui_rect.y()
                            width = ui_rect.width()
                            height = ui_rect.height()
                            print(f"Converted StandardRegion {rect_data.label}: UI({x},{y},{width},{height})")
                            return QRect(x, y, width, height)
                        else:
                            print(f"Warning: Unknown StandardRegion-like object: {type(rect_data)}")
                            return QRect()
                    elif hasattr(rect_data, 'rect') and hasattr(rect_data, 'label'):
                        # StandardRegion object - extract the UI coordinates
                        from standardized_coordinates import StandardRegion
                        if isinstance(rect_data, StandardRegion):
                            ui_rect = rect_data.rect
                            x = ui_rect.x()
                            y = ui_rect.y()
                            width = ui_rect.width()
                            height = ui_rect.height()
                            print(f"Converted StandardRegion {rect_data.label}: UI({x},{y},{width},{height})")
                            return QRect(x, y, width, height)
                        else:
                            print(f"Warning: Unknown StandardRegion-like object: {type(rect_data)}")
                            return QRect()
                    else:
                        print(f"Warning: Unknown rect type in template: {type(rect_data)}")
                        return QRect()

                # Get PDF dimensions if available
                pdf_height = None
                if self.pdf_processor and hasattr(self.pdf_processor, 'pdf_label') and self.pdf_processor.pdf_label.pixmap():
                    pdf_height = self.pdf_processor.pdf_label.pixmap().height()
                    print(f"Using PDF height for coordinate conversion: {pdf_height}")
                else:
                    print("Warning: Could not determine PDF height, using raw coordinates")

                # For single-page templates
                if template.get('template_type') == 'single':
                    # Check if we have original coordinates in the config
                    config = template.get('config', {})
                    if 'original_regions' in config and config['original_regions']:
                        print("Using original regions from config")
                        # Use original coordinates directly
                        converted_regions = {}
                        for section, rects in config['original_regions'].items():
                            converted_regions[section] = []
                            for rect_data in rects:
                                converted_rect = QRect(
                                    int(rect_data['x']),
                                    int(rect_data['y']),
                                    int(rect_data['width']),
                                    int(rect_data['height'])
                                )
                                converted_regions[section].append(converted_rect)
                                print(f"Using original region: {rect_data} -> {converted_rect}")

                        template['regions'] = converted_regions
                    # Fallback to converting serialized regions if original coordinates not available
                    elif 'regions' in template and template['regions']:
                        print("Original regions not found in config, converting serialized regions")
                        converted_regions = {}
                        for section, rects in template['regions'].items():
                            converted_regions[section] = []
                            for rect_data in rects:
                                converted_rect = convert_coordinates(rect_data, pdf_height)
                                converted_regions[section].append(converted_rect)
                                print(f"Converted region: {rect_data} -> {converted_rect}")

                        template['regions'] = converted_regions

                    # Check if we have original column lines in the config
                    config = template.get('config', {})
                    if 'original_column_lines' in config and config['original_column_lines']:
                        print("Using original column lines from config")
                        # Use original coordinates directly
                        converted_column_lines = {}
                        for section, lines in config['original_column_lines'].items():
                            converted_column_lines[section] = []
                            for line_data in lines:
                                if isinstance(line_data, list) and len(line_data) >= 2:
                                    # Check if the first two elements are dictionaries with x,y coordinates
                                    if (isinstance(line_data[0], dict) and 'x' in line_data[0] and 'y' in line_data[0] and
                                        isinstance(line_data[1], dict) and 'x' in line_data[1] and 'y' in line_data[1]):
                                        # Use original coordinates
                                        start_point = QPoint(int(line_data[0]['x']), int(line_data[0]['y']))
                                        end_point = QPoint(int(line_data[1]['x']), int(line_data[1]['y']))
                                        print(f"Using original column line coordinates from config: {start_point} -> {end_point}")

                                        # Check if we have a region index
                                        region_index = None
                                        if len(line_data) > 2 and isinstance(line_data[2], int):
                                            region_index = line_data[2]

                                        # Create line data with region index if present
                                        if region_index is not None:
                                            converted_line = [start_point, end_point, region_index]
                                        else:
                                            converted_line = [start_point, end_point]

                                        converted_column_lines[section].append(converted_line)
                                    else:
                                        print(f"Warning: Could not convert original column line from config: {line_data}")
                                else:
                                    print(f"Warning: Could not convert original column line from config: {line_data}")

                        # Store in dual coordinate format instead of legacy
                        template['drawing_column_lines'] = converted_column_lines
                    # Fallback to converting serialized column lines if original coordinates not available
                    elif template.get('drawing_column_lines') or template.get('column_lines'):
                        print("Original column lines not found in config, converting serialized column lines")
                        converted_column_lines = {}
                        # Use dual coordinate data if available, otherwise legacy
                        source_column_lines = template.get('drawing_column_lines', template.get('column_lines', {}))
                        for section, lines in source_column_lines.items():
                            converted_column_lines[section] = []
                            for line_data in lines:
                                # Check if we have original coordinates in the line data
                                if isinstance(line_data, list) and len(line_data) >= 2:
                                    # Check if the first two elements are dictionaries with x,y coordinates
                                    if (isinstance(line_data[0], dict) and 'x' in line_data[0] and 'y' in line_data[0] and
                                        isinstance(line_data[1], dict) and 'x' in line_data[1] and 'y' in line_data[1]):
                                        # Use original coordinates
                                        start_point = QPoint(int(line_data[0]['x']), int(line_data[0]['y']))
                                        end_point = QPoint(int(line_data[1]['x']), int(line_data[1]['y']))
                                        print(f"Using original column line coordinates: {start_point} -> {end_point}")

                                        # Check if we have a region index
                                        region_index = None
                                        if len(line_data) > 2 and isinstance(line_data[2], int):
                                            region_index = line_data[2]

                                        # Create line data with region index if present
                                        if region_index is not None:
                                            converted_line = [start_point, end_point, region_index]
                                        else:
                                            converted_line = [start_point, end_point]

                                        converted_column_lines[section].append(converted_line)
                                        continue

                                # Handle old format or fallback
                                # First try to extract start and end points
                                start_point = None
                                end_point = None
                                region_index = None

                                # Handle dictionary format
                                if isinstance(line_data, dict):
                                    if 'orig_start' in line_data and 'orig_end' in line_data:
                                        # Old format with original coordinates
                                        start_point_data = line_data['orig_start']
                                        end_point_data = line_data['orig_end']

                                        if isinstance(start_point_data, dict) and 'x' in start_point_data and 'y' in start_point_data:
                                            start_point = QPoint(int(start_point_data['x']), int(start_point_data['y']))

                                        if isinstance(end_point_data, dict) and 'x' in end_point_data and 'y' in end_point_data:
                                            end_point = QPoint(int(end_point_data['x']), int(end_point_data['y']))

                                        if 'region_index' in line_data:
                                            region_index = line_data['region_index']

                                # Handle list format
                                elif isinstance(line_data, list) and len(line_data) >= 2:
                                    # Try to extract points from list format
                                    if isinstance(line_data[0], dict) and 'x' in line_data[0] and 'y' in line_data[0]:
                                        start_point = QPoint(int(line_data[0]['x']), int(line_data[0]['y']))

                                    if isinstance(line_data[1], dict) and 'x' in line_data[1] and 'y' in line_data[1]:
                                        end_point = QPoint(int(line_data[1]['x']), int(line_data[1]['y']))

                                    if len(line_data) > 2 and isinstance(line_data[2], int):
                                        region_index = line_data[2]

                                # Create the converted line if we have valid points
                                if start_point and end_point:
                                    if region_index is not None:
                                        converted_column_lines[section].append([start_point, end_point, region_index])
                                    else:
                                        converted_column_lines[section].append([start_point, end_point])
                                else:
                                    print(f"Warning: Could not convert column line: {line_data}")

                        # Store in dual coordinate format instead of legacy
                        template['drawing_column_lines'] = converted_column_lines

                # For multi-page templates
                else:
                    # Check if we have page_configs with original coordinates
                    if 'page_configs' in template and template['page_configs']:
                        print("Using original coordinates from page_configs")
                        page_configs = template['page_configs']

                        # Process page regions using original coordinates from page_configs
                        if 'page_regions' in template and template['page_regions']:
                            converted_page_regions = []
                            for page_idx, page_regions in enumerate(template['page_regions']):
                                # Check if we have page_config for this page
                                if page_idx < len(page_configs) and page_configs[page_idx] and 'original_regions' in page_configs[page_idx]:
                                    # Use original coordinates from page_config
                                    original_regions = page_configs[page_idx]['original_regions']
                                    converted_regions = {}
                                    for section, rects in original_regions.items():
                                        converted_regions[section] = []
                                        for rect_data in rects:
                                            converted_rect = QRect(
                                                int(rect_data['x']),
                                                int(rect_data['y']),
                                                int(rect_data['width']),
                                                int(rect_data['height'])
                                            )
                                            converted_regions[section].append(converted_rect)
                                            print(f"Using original multi-page region from page_config: {rect_data} -> {converted_rect}")
                                    converted_page_regions.append(converted_regions)
                                else:
                                    # Fallback to converting serialized regions
                                    converted_regions = {}
                                    for section, rects in page_regions.items():
                                        converted_regions[section] = []
                                        for rect_data in rects:
                                            converted_rect = convert_coordinates(rect_data, pdf_height)
                                            converted_regions[section].append(converted_rect)
                                            print(f"Converted multi-page region: {rect_data} -> {converted_rect}")
                                    converted_page_regions.append(converted_regions)
                            template['page_regions'] = converted_page_regions

                    # Fallback to converting serialized page regions if page_configs not available
                    elif 'page_regions' in template and template['page_regions']:
                        print("Page configs not found, converting serialized page regions")
                        converted_page_regions = []
                        for page_regions in template['page_regions']:
                            converted_regions = {}
                            for section, rects in page_regions.items():
                                converted_regions[section] = []
                                for rect_data in rects:
                                    converted_rect = convert_coordinates(rect_data, pdf_height)
                                    converted_regions[section].append(converted_rect)
                                    print(f"Converted multi-page region: {rect_data} -> {converted_rect}")
                            converted_page_regions.append(converted_regions)
                        template['page_regions'] = converted_page_regions

                    # Process page column lines with coordinate conversion
                    if 'page_column_lines' in template and template['page_column_lines']:
                        # Check if we have page_configs with original coordinates
                        if 'page_configs' in template and template['page_configs']:
                            print("Using original column lines from page_configs")
                            page_configs = template['page_configs']
                            converted_page_column_lines = []

                            for page_idx, page_column_lines in enumerate(template['page_column_lines']):
                                # Check if we have page_config for this page
                                if page_idx < len(page_configs) and page_configs[page_idx] and 'original_column_lines' in page_configs[page_idx]:
                                    # Use original column lines from page_config
                                    original_column_lines = page_configs[page_idx]['original_column_lines']
                                    converted_column_lines = {}

                                    for section, lines in original_column_lines.items():
                                        converted_column_lines[section] = []
                                        for line_data in lines:
                                            if isinstance(line_data, list) and len(line_data) >= 2:
                                                # Check if the first two elements are dictionaries with x,y coordinates
                                                if (isinstance(line_data[0], dict) and 'x' in line_data[0] and 'y' in line_data[0] and
                                                    isinstance(line_data[1], dict) and 'x' in line_data[1] and 'y' in line_data[1]):
                                                    # Use original coordinates
                                                    start_point = QPoint(int(line_data[0]['x']), int(line_data[0]['y']))
                                                    end_point = QPoint(int(line_data[1]['x']), int(line_data[1]['y']))
                                                    print(f"Using original multi-page column line from page_config: {start_point} -> {end_point}")

                                                    # Check if we have a region index
                                                    region_index = None
                                                    if len(line_data) > 2 and isinstance(line_data[2], int):
                                                        region_index = line_data[2]

                                                    # Create line data with region index if present
                                                    if region_index is not None:
                                                        converted_line = [start_point, end_point, region_index]
                                                    else:
                                                        converted_line = [start_point, end_point]

                                                    converted_column_lines[section].append(converted_line)
                                                else:
                                                    print(f"Warning: Could not convert original column line from page_config: {line_data}")
                                            else:
                                                print(f"Warning: Could not convert original column line from page_config: {line_data}")

                                    converted_page_column_lines.append(converted_column_lines)
                                else:
                                    # Fallback to converting serialized column lines
                                    converted_column_lines = {}
                                    for section, lines in page_column_lines.items():
                                        converted_column_lines[section] = []
                                        for line_data in lines:
                                            # Handle new format with both original and scaled coordinates
                                            if isinstance(line_data, dict) and 'orig_start' in line_data and 'orig_end' in line_data:
                                                # Prioritize using original unscaled coordinates
                                                start_point = line_data['orig_start']
                                                end_point = line_data['orig_end']

                                                # Convert dictionary to QPoint
                                                if isinstance(start_point, dict) and 'x' in start_point and 'y' in start_point:
                                                    x = int(start_point['x'])
                                                    y = int(start_point['y'])
                                                    start_point = QPoint(x, y)

                                                if isinstance(end_point, dict) and 'x' in end_point and 'y' in end_point:
                                                    x = int(end_point['x'])
                                                    y = int(end_point['y'])
                                                    end_point = QPoint(x, y)

                                                # Create line data with region index if present
                                                if 'region_index' in line_data:
                                                    converted_column_lines[section].append([start_point, end_point, line_data['region_index']])
                                                else:
                                                    converted_column_lines[section].append([start_point, end_point])
                                            # Handle old format with list of points
                                            elif isinstance(line_data, list) or isinstance(line_data, tuple):
                                                if len(line_data) == 2:
                                                    start_point = line_data[0]
                                                    end_point = line_data[1]

                                                    # Convert dictionary to QPoint if needed
                                                    if isinstance(start_point, dict) and 'x' in start_point and 'y' in start_point:
                                                        x = int(start_point['x'])
                                                        y = int(start_point['y'])
                                                        if pdf_height is not None and template.get('uses_bottom_left', True):
                                                            y = pdf_height - y  # Convert to top-left origin
                                                        start_point = QPoint(x, y)

                                                    if isinstance(end_point, dict) and 'x' in end_point and 'y' in end_point:
                                                        x = int(end_point['x'])
                                                        y = int(end_point['y'])
                                                        if pdf_height is not None and template.get('uses_bottom_left', True):
                                                            y = pdf_height - y  # Convert to top-left origin
                                                        end_point = QPoint(x, y)

                                                    converted_column_lines[section].append([start_point, end_point])
                                                elif len(line_data) == 3:
                                                    start_point = line_data[0]
                                                    end_point = line_data[1]
                                                    rect_index = line_data[2]

                                                    # Convert dictionary to QPoint if needed
                                                    if isinstance(start_point, dict) and 'x' in start_point and 'y' in start_point:
                                                        x = int(start_point['x'])
                                                        y = int(start_point['y'])
                                                        if pdf_height is not None and template.get('uses_bottom_left', True):
                                                            y = pdf_height - y  # Convert to top-left origin
                                                        start_point = QPoint(x, y)

                                                    if isinstance(end_point, dict) and 'x' in end_point and 'y' in end_point:
                                                        x = int(end_point['x'])
                                                        y = int(end_point['y'])
                                                        if pdf_height is not None and template.get('uses_bottom_left', True):
                                                            y = pdf_height - y  # Convert to top-left origin
                                                        end_point = QPoint(x, y)

                                                    converted_column_lines[section].append([start_point, end_point, rect_index])
                                            else:
                                                # Unknown format, log error and continue
                                                print(f"Warning: Unknown column line format in multi-page template: {line_data}")
                                    converted_page_column_lines.append(converted_column_lines)
                        template['page_column_lines'] = converted_page_column_lines

                # Set configuration in PDF processor
                if self.pdf_processor:
                    # Set multi-table mode
                    if hasattr(self.pdf_processor, 'multi_table_mode'):
                        multi_table_mode = template.get('config', {}).get('multi_table_mode', False)
                        self.pdf_processor.multi_table_mode = multi_table_mode
                        print(f"Setting multi-table mode to {multi_table_mode} based on template config")

                    # Set extraction method if available
                    if hasattr(self.pdf_processor, 'current_extraction_method'):
                        # Get the latest template data from the database to ensure we have the most up-to-date extraction method
                        db_template = self.db.get_template(template_id=template_id)

                        if db_template and 'extraction_method' in db_template:
                            extraction_method = db_template['extraction_method']
                            self.pdf_processor.current_extraction_method = extraction_method
                            print(f"[DEBUG] Set extraction method from template: {extraction_method}")

                            # Update the UI dropdown if it exists
                            if hasattr(self.pdf_processor, 'extraction_method_combo'):
                                self.pdf_processor.extraction_method_combo.setCurrentText(extraction_method)
                                print(f"[DEBUG] Updated extraction method dropdown to: {extraction_method}")
                        else:
                            print(f"[DEBUG] No extraction method found in template, using default")

                    # Set extraction parameters - ALWAYS fetch from database
                    if hasattr(self.pdf_processor, 'extraction_params'):
                        # Get the latest template data from the database to ensure we have the most up-to-date extraction parameters
                        if 'db_template' not in locals():
                            db_template = self.db.get_template(template_id=template_id)

                        # Print the entire db_template for debugging
                        print(f"[DEBUG] Database template: {db_template}")

                        if db_template and 'config' in db_template and 'extraction_params' in db_template['config']:
                            # Use extraction parameters from database WITHOUT adding any defaults
                            extraction_params = db_template['config']['extraction_params']
                            print(f"[DEBUG] Setting extraction parameters from database exactly as stored: {extraction_params}")

                            # Verify extraction parameters structure
                            if not isinstance(extraction_params, dict):
                                print(f"[WARNING] Extraction parameters from database are not a dictionary: {type(extraction_params)}")
                                extraction_params = {}

                            # Ensure all section parameters exist but don't add default values
                            for section in ['header', 'items', 'summary']:
                                if section not in extraction_params:
                                    extraction_params[section] = {}
                                    print(f"[DEBUG] Added empty section {section} without defaults")

                                # Ensure section parameters have the required structure
                                section_params = extraction_params[section]

                                # Get global parameters
                                global_flavor = extraction_params.get('flavor', 'stream')
                                global_split_text = extraction_params.get('split_text', True)
                                global_strip_text = extraction_params.get('strip_text', '\n')

                                # Add missing parameters to section if they don't exist
                                if 'flavor' not in section_params:
                                    section_params['flavor'] = global_flavor
                                    print(f"[DEBUG] Added flavor={global_flavor} to {section} section")

                                if 'split_text' not in section_params:
                                    section_params['split_text'] = global_split_text
                                    print(f"[DEBUG] Added split_text={global_split_text} to {section} section")

                                if 'strip_text' not in section_params:
                                    section_params['strip_text'] = global_strip_text
                                    print(f"[DEBUG] Added strip_text={global_strip_text} to {section} section")

                                if 'edge_tol' not in section_params:
                                    section_params['edge_tol'] = 0.5
                                    print(f"[DEBUG] Added edge_tol=0.5 to {section} section")

                            # Set the extraction parameters exactly as they are in the database
                            self.pdf_processor.extraction_params = extraction_params

                            # Print the extraction parameters that were set
                            print(f"[DEBUG] Final extraction parameters set: {self.pdf_processor.extraction_params}")

                            # Verify row_tol values
                            for section in ['header', 'items', 'summary']:
                                if 'row_tol' in extraction_params.get(section, {}):
                                    print(f"[DEBUG] {section} row_tol: {extraction_params[section]['row_tol']}")
                                else:
                                    print(f"[DEBUG] {section} row_tol not found in extraction parameters")

                        elif 'config' in template and 'extraction_params' in template['config']:
                            # Fallback to template config if database doesn't have extraction parameters
                            # Use exactly as stored without adding defaults
                            extraction_params = template['config']['extraction_params']
                            print(f"[DEBUG] Setting extraction parameters from template config exactly as stored: {extraction_params}")

                            # Verify extraction parameters structure
                            if not isinstance(extraction_params, dict):
                                print(f"[WARNING] Extraction parameters from template are not a dictionary: {type(extraction_params)}")
                                extraction_params = {}

                            # Ensure all section parameters exist but don't add default values
                            for section in ['header', 'items', 'summary']:
                                if section not in extraction_params:
                                    extraction_params[section] = {}
                                    print(f"[DEBUG] Added empty section {section} without defaults")

                                # Ensure section parameters have the required structure
                                section_params = extraction_params[section]

                                # Get global parameters
                                global_flavor = extraction_params.get('flavor', 'stream')
                                global_split_text = extraction_params.get('split_text', True)
                                global_strip_text = extraction_params.get('strip_text', '\n')

                                # Add missing parameters to section if they don't exist
                                if 'flavor' not in section_params:
                                    section_params['flavor'] = global_flavor
                                    print(f"[DEBUG] Added flavor={global_flavor} to {section} section")

                                if 'split_text' not in section_params:
                                    section_params['split_text'] = global_split_text
                                    print(f"[DEBUG] Added split_text={global_split_text} to {section} section")

                                if 'strip_text' not in section_params:
                                    section_params['strip_text'] = global_strip_text
                                    print(f"[DEBUG] Added strip_text={global_strip_text} to {section} section")

                                if 'edge_tol' not in section_params:
                                    section_params['edge_tol'] = 0.5
                                    print(f"[DEBUG] Added edge_tol=0.5 to {section} section")

                            # Set the extraction parameters exactly as they are in the template
                            self.pdf_processor.extraction_params = extraction_params

                            # Print the extraction parameters that were set
                            print(f"[DEBUG] Final extraction parameters set: {self.pdf_processor.extraction_params}")

                            # Verify row_tol values
                            for section in ['header', 'items', 'summary']:
                                if 'row_tol' in extraction_params.get(section, {}):
                                    print(f"[DEBUG] {section} row_tol: {extraction_params[section]['row_tol']}")
                                else:
                                    print(f"[DEBUG] {section} row_tol not found in extraction parameters")

                        else:
                            # If no extraction parameters found, initialize with empty structure
                            # Do not add any default values
                            self.pdf_processor.extraction_params = {
                                'header': {},
                                'items': {},
                                'summary': {},
                                'flavor': 'stream'  # Only add flavor as it's required for extraction
                            }
                            print(f"[DEBUG] No extraction parameters found, initializing with empty structure without defaults")

                            # Print the extraction parameters that were set
                            print(f"[DEBUG] Final extraction parameters set: {self.pdf_processor.extraction_params}")

                    # Set the entire config object for reference - always use the latest from database
                    if hasattr(self.pdf_processor, 'template_config'):
                        # Use the db_template we already fetched above if available
                        if 'db_template' in locals() and db_template and 'config' in db_template:
                            self.pdf_processor.template_config = db_template['config']
                            print(f"[DEBUG] Setting complete template config from database")
                        else:
                            # Fallback to template config if database fetch failed
                            self.pdf_processor.template_config = template.get('config', {})
                            print(f"[DEBUG] Setting complete template config from template (fallback)")

                # Set up table_areas dictionary in the pdf_processor for structured storage
                if self.pdf_processor:
                    template_type = template.get('template_type', 'single')
                    if template_type == 'single' and 'regions' in template:
                        self.pdf_processor.table_areas = {}

                        # Build table_areas for each section - handle standard format
                        for section, region_list in template['regions'].items():
                            for i, region_item in enumerate(region_list):
                                table_label = f"{section}_table_{i+1}"

                                # Extract coordinates from clean dual coordinates format
                                try:
                                    from clean_region_utils import get_drawing_coordinates, get_region_name
                                    rect = get_drawing_coordinates(region_item)
                                    name = get_region_name(region_item)
                                except ImportError:
                                    # Fallback for when clean_region_utils is not available
                                    print(f"[WARNING] clean_region_utils not available, using fallback for {section}[{i}]")
                                    # Handle StandardRegion objects from database
                                    if hasattr(region_item, 'rect') and hasattr(region_item, 'label'):
                                        rect = region_item.rect
                                        name = region_item.label
                                    elif isinstance(region_item, dict) and 'rect' in region_item:
                                        rect = region_item['rect']
                                        name = region_item.get('label', f"{section}_{i}")
                                    else:
                                        print(f"Skipping invalid region in {section}[{i}]: unsupported format")
                                        continue
                                except ValueError as e:
                                    print(f"Skipping invalid region in {section}[{i}]: {e}")
                                    continue

                                # Get columns for this table - use dual coordinate data if available
                                columns = []
                                # Check dual coordinate column lines first, then legacy
                                drawing_column_lines = template.get('drawing_column_lines', template.get('column_lines', {}))
                                if section in drawing_column_lines:
                                    for line_data in drawing_column_lines[section]:
                                        if len(line_data) >= 2:
                                            if len(line_data) == 2 or (len(line_data) == 3 and line_data[2] == i):
                                                # This line belongs to the current table
                                                columns.append(line_data[0].x())

                                # Create the table area entry
                                self.pdf_processor.table_areas[table_label] = {
                                    'type': section,
                                    'index': i,
                                    'rect': rect,
                                    'name': name,
                                    'columns': sorted(columns) if columns else []
                                }
                                print(f"Created table_area: {table_label} with {len(columns)} columns, name='{name}'")

                    # For multi-page templates, table_areas will be set up when pages are displayed

                # Check for YAML/JSON template - always use the latest from database
                yaml_template = None

                # First try to get the template from the database
                if 'db_template' in locals() and db_template and 'json_template' in db_template:
                    yaml_template = db_template['json_template']
                    print(f"\n[DEBUG] YAML template found in database: {yaml_template is not None}")
                    template_source = "database"
                # Fallback to template data if database doesn't have it
                elif 'json_template' in template:
                    yaml_template = template['json_template']
                    print(f"\n[DEBUG] YAML template found in template data: {yaml_template is not None}")
                    template_source = "template data"

                if yaml_template:
                    print(f"[DEBUG] YAML template type: {type(yaml_template)}")
                    if isinstance(yaml_template, dict):
                        print(f"[DEBUG] YAML template keys: {list(yaml_template.keys())}")

                        # Format as YAML (preferred)
                        try:
                            import yaml
                            formatted_yaml = yaml.dump(yaml_template, default_flow_style=False, allow_unicode=True)
                            print(f"[DEBUG] YAML template preview from {template_source} (first 200 chars): {formatted_yaml[:200]}...")

                            # Set the YAML template in the PDF processor if it has the attribute
                            if self.pdf_processor and hasattr(self.pdf_processor, 'template_preview'):
                                print(f"[DEBUG] Setting YAML template in PDF processor's template_preview")
                                self.pdf_processor.template_preview.setText(formatted_yaml)
                        except Exception as e:
                            # Fallback to JSON if YAML formatting fails
                            print(f"[DEBUG] Failed to format as YAML, using JSON: {str(e)}")
                            formatted_json = json.dumps(yaml_template, indent=2)
                            print(f"[DEBUG] JSON template preview from {template_source} (first 200 chars): {formatted_json[:200]}...")

                            # Set the JSON template in the PDF processor if it has the attribute
                            if self.pdf_processor and hasattr(self.pdf_processor, 'template_preview'):
                                print(f"[DEBUG] Setting JSON template in PDF processor's template_preview")
                                self.pdf_processor.template_preview.setText(formatted_json)
                else:
                    print(f"[DEBUG] YAML/JSON template is None or empty")

                # Add debugging information
                print("\n[DEBUG] Template data that will be applied:")
                if template.get('template_type') == 'single':
                    # Use dual coordinate regions if available, otherwise use legacy regions
                    regions_data = template.get('drawing_regions', template.get('regions', {}))
                    for section, rects in regions_data.items():
                        print(f"Section: {section} - {len(rects)} regions")
                        for i, rect_item in enumerate(rects):
                            # Handle clean dual coordinates format
                            try:
                                from clean_region_utils import get_drawing_coordinates, get_region_name
                                rect = get_drawing_coordinates(rect_item)
                                name = get_region_name(rect_item)
                                print(f"  Region {i}: x={rect.x()}, y={rect.y()}, width={rect.width()}, height={rect.height()}, name='{name}'")
                            except ImportError:
                                # Fallback for when clean_region_utils is not available
                                print(f"  Region {i}: clean_region_utils not available, using fallback")
                                # Handle StandardRegion objects from database
                                if hasattr(rect_item, 'rect') and hasattr(rect_item, 'label'):
                                    rect = rect_item.rect
                                    name = rect_item.label
                                    print(f"  Region {i}: x={rect.x()}, y={rect.y()}, width={rect.width()}, height={rect.height()}, name='{name}'")
                                elif isinstance(rect_item, dict) and 'rect' in rect_item:
                                    rect = rect_item['rect']
                                    name = rect_item.get('label', f"{section}_{i}")
                                    print(f"  Region {i}: x={rect.x()}, y={rect.y()}, width={rect.width()}, height={rect.height()}, name='{name}'")
                                else:
                                    print(f"  Region {i}: Unable to parse region format")
                            except ValueError:
                                # Fallback for any format issues
                                print(f"  Region {i}: Unable to parse region format")

                    # Use dual coordinate column lines if available, otherwise use legacy column lines
                    column_lines_data = template.get('drawing_column_lines', template.get('column_lines', {}))
                    for section, lines in column_lines_data.items():
                        print(f"Column lines for section: {section} - {len(lines)} lines")
                        for i, line in enumerate(lines):
                            if hasattr(line, 'drawing_start_x'):  # DualCoordinateColumnLine
                                print(f"  Line {i}: start=({line.drawing_start_x}, {line.drawing_start_y}), end=({line.drawing_end_x}, {line.drawing_end_y})")
                            elif len(line) == 2:  # Legacy format
                                print(f"  Line {i}: start=({line[0].x()}, {line[0].y()}), end=({line[1].x()}, {line[1].y()})")
                            elif len(line) == 3:  # Legacy format with rect_index
                                print(f"  Line {i}: start=({line[0].x()}, {line[0].y()}), end=({line[1].x()}, {line[1].y()}), rect_index={line[2]}")
                else:
                    print(f"Multi-page template with {template.get('page_count', 1)} pages")
                    for page_idx, page_regions in enumerate(template.get('page_regions', [])):
                        print(f"Page {page_idx + 1} regions:")
                        for section, rects in page_regions.items():
                            print(f"  Section: {section} - {len(rects)} regions")
                            for i, rect_item in enumerate(rects):
                                # Handle clean dual coordinates format
                                try:
                                    from clean_region_utils import get_drawing_coordinates, get_region_name
                                    rect = get_drawing_coordinates(rect_item)
                                    name = get_region_name(rect_item)
                                    print(f"    Region {i}: x={rect.x()}, y={rect.y()}, width={rect.width()}, height={rect.height()}, name='{name}'")
                                except ImportError:
                                    # Fallback for when clean_region_utils is not available
                                    print(f"    Region {i}: clean_region_utils not available, using fallback")
                                    # Handle StandardRegion objects from database
                                    if hasattr(rect_item, 'rect') and hasattr(rect_item, 'label'):
                                        rect = rect_item.rect
                                        name = rect_item.label
                                        print(f"    Region {i}: x={rect.x()}, y={rect.y()}, width={rect.width()}, height={rect.height()}, name='{name}'")
                                    elif isinstance(rect_item, dict) and 'rect' in rect_item:
                                        rect = rect_item['rect']
                                        name = rect_item.get('label', f"{section}_{i}")
                                        print(f"    Region {i}: x={rect.x()}, y={rect.y()}, width={rect.width()}, height={rect.height()}, name='{name}'")
                                    else:
                                        print(f"    Region {i}: Unable to parse region format")
                                except ValueError:
                                    # Fallback for any format issues
                                    print(f"    Region {i}: Unable to parse region format")



                # Ask user to select a PDF file
                file_path, _ = QFileDialog.getOpenFileName(
                    self, "Select PDF File to Apply Template", "", "PDF Files (*.pdf)"
                )

                if not file_path:
                    # User cancelled, don't apply template
                    QMessageBox.information(
                        self,
                        "Template Application Cancelled",
                        "You need to select a PDF file to apply the template."
                    )
                    return

                # Add the file path to the template data
                template['selected_pdf_path'] = file_path
                print(f"\n[DEBUG] Selected PDF file: {file_path}")
                print(f"[DEBUG] Template data: {template['name']}, type: {template.get('template_type', 'single')}")
                # Use dual coordinate data for counts
                drawing_regions = template.get('drawing_regions', template.get('regions', {}))
                drawing_column_lines = template.get('drawing_column_lines', template.get('column_lines', {}))
                print(f"[DEBUG] Regions count: {len(drawing_regions)}")
                print(f"[DEBUG] Column lines count: {len(drawing_column_lines)}")
                if template.get('template_type') == 'multi':
                    print(f"[DEBUG] Page regions count: {len(template.get('page_regions', []))}")
                    print(f"[DEBUG] Page column lines count: {len(template.get('page_column_lines', []))}")
                    print(f"[DEBUG] Page configs count: {len(template.get('page_configs', []))}")
                print(f"[DEBUG] Config: {template.get('config', {})}")

                # Make sure multi-page settings are properly included
                if template.get('template_type') == 'multi':
                    # Set is_multi_page flag
                    template['is_multi_page'] = True

                    # Multi-page options removed - using simplified page-wise approach

                print("[DEBUG] Emitting template_selected signal...")


                # Ensure page_configs are properly included in the template data
                if template.get('template_type') == 'multi' and 'page_configs' in template:
                    print(f"[DEBUG] Ensuring page_configs are properly included: {len(template['page_configs'])} configs")
                    # Make sure page_configs are properly formatted and accessible
                    for i, page_config in enumerate(template['page_configs']):
                        if page_config and 'original_regions' in page_config:
                            print(f"[DEBUG] Page {i+1} has original_regions with {len(page_config['original_regions'])} sections")

                # Emit the template selection signal with the template data
                # The PDFProcessor will receive this data and apply it to the PDF
                # The PDFLabel will handle the scaling when drawing the regions
                # Add a flag to force extraction to ensure results are updated immediately
                template['force_extraction'] = True
                self.template_selected.emit(template)

                # Show success message - commented out as requested
                # success_msg = QMessageBox(self)
                # success_msg.setWindowTitle("Template Applied")
                # success_msg.setText("Template Applied Successfully")
                # success_msg.setInformativeText(f"The template '{template['name']}' has been applied to '{os.path.basename(file_path)}'. You can now use it for PDF processing.")
                # success_msg.setIcon(QMessageBox.Information)
                # success_msg.setStyleSheet("QLabel { color: black; }")
                # success_msg.exec()

        except Exception as e:
            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Error Applying Template")
            error_dialog.setInformativeText(f"An error occurred: {str(e)}")
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setStyleSheet("QLabel { color: black; }")
            error_dialog.exec()
            # Print detailed error information to help with debugging
            import traceback
            traceback.print_exc()

    def closeEvent(self, event):
        """Close database connection when widget is closed"""
        self.db.close()
        super().closeEvent(event)

    def refresh(self):
        """Refresh the template list and update the UI"""
        # Reload templates from the database
        self.load_templates()

        # Update any UI components that need refreshing
        print("Template manager refreshed")

    def reset_database(self):
        """Delete and recreate the SQLite database, removing all templates"""
        # Confirm with the user before proceeding
        response = QMessageBox.warning(
            self,
            "Reset Database",
            "This will delete ALL templates from the database.\n\n"
            "This action CANNOT be undone. Are you sure you want to continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if response == QMessageBox.No:
            return

        try:
            # Close all existing database connections
            self.close_all_connections()

            # Add a small delay to ensure connections are fully closed
            import time
            time.sleep(1)

            # Delete the database file with retry mechanism
            max_retries = 5
            retry_delay = 1  # seconds
            db_path = 'invoice_templates.db'

            for attempt in range(max_retries):
                try:
                    if os.path.exists(db_path):
                        # Force close any remaining handles (Windows specific)
                        if os.name == 'nt':
                            import ctypes
                            kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)

                            try:
                                # Try to take ownership of the file
                                handle = kernel32.CreateFileW(
                                    db_path,
                                    0x80000000 | 0x40000000,  # GENERIC_READ | GENERIC_WRITE
                                    0,  # No sharing
                                    None,  # No security
                                    3,  # OPEN_EXISTING
                                    0x80,  # FILE_ATTRIBUTE_NORMAL
                                    None  # No template
                                )

                                if handle != -1:  # If successful
                                    kernel32.CloseHandle(handle)
                            except Exception as e:
                                print(f"Warning: Could not take file ownership: {str(e)}")

                        # Try to remove the file
                        os.remove(db_path)
                        print("Successfully deleted database file")
                        break
                    else:
                        print("Database file doesn't exist, creating new one")
                        break

                except PermissionError as e:
                    if attempt < max_retries - 1:
                        print(f"Attempt {attempt + 1} failed, retrying in {retry_delay} seconds...")
                        time.sleep(retry_delay)
                        continue
                    else:
                        raise Exception("Maximum retries reached. Could not delete database file.")

            # Create a new database with retry mechanism
            for attempt in range(max_retries):
                try:
                    self.db = InvoiceDatabase()  # This will recreate the database
                    break
                except Exception as e:
                    if attempt < max_retries - 1:
                        print(f"Failed to create new database, attempt {attempt + 1}. Retrying...")
                        time.sleep(retry_delay)
                        continue
                    else:
                        raise Exception("Failed to create new database after multiple attempts")

            # Refresh the templates table
            self.templates_table.setRowCount(0)

            QMessageBox.information(
                self,
                "Database Reset",
                "The database has been successfully reset. All templates have been removed."
            )

            print("\n" + "="*80)
            print("DATABASE RESET SUCCESSFUL")
            print("All templates have been removed and the database has been recreated.")
            print("="*80)

        except Exception as e:
            error_msg = (
                f"Failed to reset database: {str(e)}\n\n"
                "Please try the following:\n"
                "1. Close all other applications\n"
                "2. Wait a few seconds\n"
                "3. Try again\n\n"
                "If the problem persists, restart the application."
            )
            QMessageBox.critical(self, "Error", error_msg)
            print(f"Error resetting database: {str(e)}")
            import traceback
            traceback.print_exc()

    def close_all_connections(self):
        """Close all database connections"""
        try:
            # Close the main database connection
            if hasattr(self, 'db'):
                self.db.close()
                delattr(self, 'db')

            # Force Python's garbage collection to clean up any lingering connections
            import gc
            gc.collect()

            # Close global database instance
            try:
                from unified_invoice_viewer import DatabaseManager
                DatabaseManager.close()
            except Exception as e:
                print(f"Warning: Could not close DatabaseManager: {str(e)}")

            # Close bulk processor connections
            try:
                from bulk_processor import BulkProcessor
                if hasattr(BulkProcessor, 'conn'):
                    BulkProcessor.conn.close()
                    BulkProcessor.conn = None
            except Exception as e:
                print(f"Warning: Could not close BulkProcessor connection: {str(e)}")

            # Close any SQLite connections in memory
            try:
                import sqlite3
                for obj in gc.get_objects():
                    if isinstance(obj, sqlite3.Connection):
                        try:
                            obj.close()
                        except Exception:
                            pass
            except Exception as e:
                print(f"Warning: Error during SQLite connections cleanup: {str(e)}")

            print("Successfully closed all database connections")

        except Exception as e:
            print(f"Error in close_all_connections: {str(e)}")
            import traceback
            traceback.print_exc()  # Add this method to the TemplateManager class
    def validate_template_data(self, template_data):
        """Validate template data before saving to database"""
        if not template_data:
            raise ValueError("Template data is empty")

        # Check basic fields that are always required
        required_fields = ['name', 'template_type']
        for field in required_fields:
            if field not in template_data:
                raise ValueError(f"Missing required field: {field}")

        # Check type-specific required fields
        if template_data['template_type'] == 'multi':
            # For multi-page templates
            if 'page_count' not in template_data or not template_data['page_count']:
                raise ValueError("Multi-page templates must have a page count")

            if 'page_regions' not in template_data or not template_data['page_regions']:
                raise ValueError("Multi-page templates must have page regions defined")
        else:
            # For single-page templates - check for dual coordinate regions
            if 'drawing_regions' not in template_data and 'extraction_regions' not in template_data:
                raise ValueError("Single-page templates must have dual coordinate regions defined")

            # Validate dual coordinate regions data
            drawing_regions = template_data.get('drawing_regions', {})
            if not any(drawing_regions.values()) if drawing_regions else True:
                raise ValueError("No dual coordinate regions defined in template")

        return True

    def get_template_page_for_pdf_page(self, pdf_page_index, pdf_total_pages, template_data):
        """Determine which template page to use for a given PDF page

        Args:
            pdf_page_index (int): The 0-based index of the PDF page
            pdf_total_pages (int): The total number of pages in the PDF
            template_data (dict): The template data dictionary

        Returns:
            int: The 0-based index of the template page to use
        """
        # Get template type and page count
        template_type = template_data.get('template_type', 'single')
        template_page_count = template_data.get('page_count', 1)

        # For single-page templates, always use the first (and only) page
        if template_type == 'single':
            return 0

        # Complex page mapping logic removed - using simplified page-wise approach
        # Page mapping features will be implemented later as a separate enhancement

        # For now, use simple page mapping: use the exact page index, but don't exceed template page count
        return min(pdf_page_index, template_page_count - 1)

    # Complex template application method removed - using simplified page-wise approach
    # Page mapping features will be implemented later as a separate enhancement

    # Removed extract_with_new_params function as it is no longer needed

    def add_template(self):
        """Add a new template using factory pattern"""
        try:
            # Show dialog to get template name and description
            dialog = SaveTemplateDialog(self)
            if dialog.exec() == QDialog.Accepted:
                template_data = dialog.get_template_data()
                name = template_data["name"]
                description = template_data["description"]

                # Validate template name using factory
                if not ValidationFactory.validate_template_name(name):
                    UIMessageFactory.invalid_name_error(self)
                    return

                # Sanitize the name
                name = ValidationFactory.sanitize_template_name(name)

                # Create database operation factory
                db_factory = get_database_factory(self.db)

                # Try to save template using factory
                template_id = db_factory.save_template_safe(
                    name=name,
                    description=description,
                    template_type="single",
                    currency="INR"
                )

                if template_id is None:
                    # Template already exists
                    UIMessageFactory.template_exists_error(self, name)
                    return

                # Success - template created
                print(f"Created new template '{name}' with ID: {template_id}")

                # Show success message using factory
                UIMessageFactory.template_saved_success(self, name)

                # Refresh the template list
                self.load_templates()

                # Open the template for editing
                self.edit_template(template_id)

        except Exception as e:
            # Use factory for error handling
            UIMessageFactory.show_error(
                self,
                "Error Creating Template",
                f"An error occurred while trying to create the template:\n\n{str(e)}\n\nPlease try again or contact support if this issue persists."
            )

            print(f"Error in add_template: {str(e)}")
            import traceback
            traceback.print_exc()

    def export_template(self):
        """Export a template to a JSON file"""
        try:
            # Get selected template
            selected_rows = self.templates_table.selectionModel().selectedRows()
            if not selected_rows:
                # No row selected, show a dialog to select a template
                templates = self.db.get_all_templates()
                if not templates:
                    QMessageBox.warning(
                        self,
                        "No Templates",
                        "There are no templates to export.",
                        QMessageBox.Ok
                    )
                    return

                # Create a dialog to select a template
                dialog = QDialog(self)
                dialog.setWindowTitle("Select Template to Export")
                dialog.setMinimumWidth(400)

                layout = QVBoxLayout(dialog)

                # Add explanation
                explanation = QLabel("Select a template to export:")
                explanation.setWordWrap(True)
                layout.addWidget(explanation)

                # Create a list widget for template selection
                template_list = QListWidget()
                for template in templates:
                    item = QListWidgetItem(f"{template['name']} ({template['template_type']})")
                    item.setData(Qt.UserRole, template['id'])
                    template_list.addItem(item)

                layout.addWidget(template_list)

                # Add buttons
                buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                buttons.accepted.connect(dialog.accept)
                buttons.rejected.connect(dialog.reject)
                layout.addWidget(buttons)

                # Show dialog and get result
                if dialog.exec() == QDialog.Accepted:
                    # Get selected template
                    selected_items = template_list.selectedItems()
                    if not selected_items:
                        QMessageBox.warning(
                            self,
                            "No Template Selected",
                            "Please select a template to export.",
                            QMessageBox.Ok
                        )
                        return

                    template_id = selected_items[0].data(Qt.UserRole)
                else:
                    return  # User cancelled
            else:
                # Get template ID from selected row
                row = selected_rows[0].row()
                template_id = self.get_template_id_from_row(row)

            # Get template data
            template = self.db.get_template(template_id=template_id)
            if not template:
                QMessageBox.warning(
                    self,
                    "Template Not Found",
                    "The selected template could not be found.",
                    QMessageBox.Ok
                )
                return

            # Create a file dialog to save the template
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Template",
                f"{template['name'].replace(' ', '_')}.json",
                "JSON Files (*.json)"
            )

            if not file_path:
                return  # User cancelled

            # Ensure file has .json extension
            if not file_path.lower().endswith('.json'):
                file_path += '.json'

            # Create a copy of the template data without the ID
            export_data = template.copy()
            if 'id' in export_data:
                del export_data['id']

            # Add export metadata
            export_data['export_date'] = datetime.now().isoformat()
            export_data['export_version'] = '1.0'

            # Write the template to the file
            with open(file_path, 'w') as f:
                json.dump(export_data, f, indent=2)

            # Show success message using factory
            UIMessageFactory.show_info(
                self,
                "Template Exported",
                f"Template '{template['name']}' has been exported to {file_path}."
            )

        except Exception as e:
            # Use factory for error handling
            UIMessageFactory.show_error(
                self,
                "Error Exporting Template",
                f"An error occurred while trying to export the template:\n\n{str(e)}\n\nPlease try again or contact support if this issue persists."
            )

            print(f"Error in export_template: {str(e)}")
            import traceback
            traceback.print_exc()

    def import_template(self):
        """Import a template from a JSON file"""
        try:
            # Create a file dialog to select the template file
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Import Template",
                "",
                "JSON Files (*.json)"
            )

            if not file_path:
                return  # User cancelled

            # Read the template from the file
            with open(file_path, 'r') as f:
                import_data = json.load(f)

            # Validate the imported data
            if not isinstance(import_data, dict):
                QMessageBox.warning(
                    self,
                    "Invalid Template",
                    "The selected file does not contain a valid template.",
                    QMessageBox.Ok
                )
                return

            # Check required fields
            required_fields = ['name', 'template_type', 'regions', 'column_lines', 'config']
            for field in required_fields:
                if field not in import_data:
                    QMessageBox.warning(
                        self,
                        "Invalid Template",
                        f"The template is missing the required field: {field}",
                        QMessageBox.Ok
                    )
                    return

            # Check if a template with the same name already exists
            existing_template = self.db.get_template(template_name=import_data['name'])
            if existing_template:
                # Ask user if they want to overwrite
                reply = QMessageBox.question(
                    self,
                    "Template Already Exists",
                    f"A template with the name '{import_data['name']}' already exists. Do you want to overwrite it?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )

                if reply == QMessageBox.Yes:
                    # Delete existing template
                    self.db.delete_template(template_id=existing_template['id'])
                else:
                    # Ask for a new name
                    new_name, ok = QInputDialog.getText(
                        self,
                        "New Template Name",
                        "Enter a new name for the imported template:",
                        text=import_data['name'] + " (imported)"
                    )

                    if not ok or not new_name.strip():
                        return  # User cancelled

                    import_data['name'] = new_name.strip()

            # Create a progress dialog
            progress = QProgressDialog("Importing template...", None, 0, 100, self)
            progress.setWindowTitle("Importing Template")
            progress.setWindowModality(Qt.WindowModal)
            progress.setMinimumDuration(0)  # Show immediately
            progress.setValue(0)
            progress.setAutoClose(True)
            progress.setAutoReset(True)
            progress.setCancelButton(None)  # No cancel button
            progress.setFixedSize(300, 100)
            progress.setStyleSheet("QLabel { color: black; }")
            progress.show()

            # Process events to ensure the dialog is displayed
            QApplication.processEvents()

            try:
                # Update progress
                progress.setValue(30)
                QApplication.processEvents()

                # Save the template using factory pattern
                db_factory = get_database_factory(self.db)

                # For multi-page templates, we need to use direct database call
                # as the factory doesn't support multi-page templates yet
                if import_data['template_type'] == 'multi':
                    # For multi-page templates - use direct database call
                    template_id = self.db.save_template(
                        name=import_data['name'],
                        description=import_data.get('description', ''),
                        regions=import_data['regions'],
                        column_lines=import_data['column_lines'],
                        config=import_data['config'],
                        template_type=import_data['template_type'],
                        page_count=import_data.get('page_count', 1),
                        page_regions=import_data.get('page_regions', []),
                        page_column_lines=import_data.get('page_column_lines', []),
                        page_configs=import_data['config'].get('page_configs', []),
                        json_template=import_data.get('json_template')
                    )
                else:
                    # For single-page templates - use factory
                    template_id = db_factory.save_template_safe(
                        name=import_data['name'],
                        description=import_data.get('description', ''),
                        template_type=import_data['template_type'],
                        regions=import_data['regions'],
                        column_lines=import_data['column_lines'],
                        config=import_data['config'],
                        json_template=import_data.get('json_template')
                    )

                    if template_id is None:
                        raise Exception("Template with this name already exists")

                # Update progress
                progress.setValue(70)
                QApplication.processEvents()

                # Refresh the template list
                self.load_templates()

                # Update progress to completion
                progress.setValue(100)
                QApplication.processEvents()

                # Show success message using factory
                UIMessageFactory.show_info(
                    self,
                    "Template Imported",
                    f"Template '{import_data['name']}' has been imported successfully."
                )

            except Exception as inner_e:
                # Close the progress dialog
                progress.close()
                QApplication.processEvents()

                error_message = f"""
<h3>Error Importing Template</h3>
<p>An error occurred while trying to import the template:</p>
<p style='color: #D32F2F;'>{str(inner_e)}</p>
<p>The template may not have been imported properly.</p>
"""
                error_dialog = QMessageBox(self)
                error_dialog.setWindowTitle("Import Error")
                error_dialog.setText("Error Importing Template")
                error_dialog.setInformativeText(error_message)
                error_dialog.setIcon(QMessageBox.Critical)
                error_dialog.setStyleSheet("QLabel { color: black; }")
                error_dialog.exec()

                print(f"Error in import_template save operation: {str(inner_e)}")
                import traceback
                traceback.print_exc()

            finally:
                # Make sure progress dialog is closed
                try:
                    progress.close()
                    QApplication.processEvents()
                except Exception as close_e:
                    print(f"Error closing progress dialog: {str(close_e)}")

        except json.JSONDecodeError as json_e:
            QMessageBox.warning(
                self,
                "Invalid JSON",
                f"The selected file contains invalid JSON: {str(json_e)}",
                QMessageBox.Ok
            )
        except Exception as e:
            error_message = f"""
<h3>Error Importing Template</h3>
<p>An error occurred while trying to import the template:</p>
<p style='color: #D32F2F;'>{str(e)}</p>
<p>Please try again or contact support if this issue persists.</p>
"""
            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Error Importing Template")
            error_dialog.setInformativeText(error_message)
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setStyleSheet("QLabel { color: black; }")
            error_dialog.exec()

            print(f"Error in import_template: {str(e)}")
            import traceback
            traceback.print_exc()
