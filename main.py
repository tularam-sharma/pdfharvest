import sys
import os
import traceback
import argparse
import logging
from datetime import datetime

# Import factory modules for code deduplication
from common_factories import (
    TemplateFactory, DatabaseOperationFactory, UIMessageFactory,
    ValidationFactory, get_database_factory
)
from ui_component_factory import UIComponentFactory, LayoutFactory
from simplified_extraction_engine import get_extraction_engine

# Set up logging - default level will be updated based on command line args
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("PDFHarvest")

# Only import essential Qt components for startup
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox
from PySide6.QtCore import Qt

# Lazy import flags to track what's been loaded
_qt_widgets_loaded = False
_heavy_modules_loaded = False

def _load_qt_widgets():
    """Lazy load Qt widgets when needed"""
    global _qt_widgets_loaded
    if not _qt_widgets_loaded:
        global QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QStackedWidget
        global QFileDialog, QScrollArea, QFrame, QSplitter, QGridLayout, QLineEdit
        global QComboBox, QListWidget, QProgressBar, QTableWidget, QTableWidgetItem
        global QHeaderView, QGroupBox, QMenu, QMenuBar, Signal, QObject, QRect
        global QFont, QIcon, QAction

        from PySide6.QtWidgets import (
            QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QStackedWidget,
            QFileDialog, QScrollArea, QFrame, QSplitter, QGridLayout, QLineEdit,
            QComboBox, QListWidget, QProgressBar, QTableWidget, QTableWidgetItem,
            QHeaderView, QGroupBox, QMenu, QMenuBar
        )
        from PySide6.QtCore import Signal, QObject, QRect
        from PySide6.QtGui import QFont, QIcon, QAction
        _qt_widgets_loaded = True

def _load_heavy_modules():
    """Lazy load heavy modules when needed"""
    global _heavy_modules_loaded
    if not _heavy_modules_loaded:
        global TemplateManager, BulkProcessor, UserManagement, SplitScreenInvoiceProcessor
        global MainDashboard, TemplateManagementCard, BulkExtractionCard, UploadProcessCard, RoleBasedWidget
        global UserManagementDialog, RoleManagementDialog
        global check_activation, show_license_manager_dialog
        global initialize_database_protection, cleanup_database_protection

        from template_manager import TemplateManager
        from bulk_processor import BulkProcessor
        from user_management import UserManagement
        from split_screen_invoice_processor import SplitScreenInvoiceProcessor
        from role_based_ui import (
            MainDashboard, TemplateManagementCard,
            BulkExtractionCard, UploadProcessCard, RoleBasedWidget
        )
        from user_management_ui import UserManagementDialog, RoleManagementDialog
        from activation_dialog import check_activation, show_license_manager_dialog
        from db_protection import initialize_database_protection, cleanup_database_protection
        _heavy_modules_loaded = True

# Define MultiPageProcessor as an alias - will be set after lazy loading
MultiPageProcessor = None

# Import CLI functionality
try:
    from tests.pdf_extractor_cli import process_pdf_folder
except ImportError:
    try:
        from pdf_extractor_cli import process_pdf_folder
    except ImportError:
        def process_pdf_folder(*args, **kwargs):
            print("Error: CLI module not found")
            return False

class PDFHarvest(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Harvest")
        self.setMinimumSize(1200, 800)

        # Track window state globally
        self.is_fullscreen = False
        self.is_maximized = False

        # Load Qt widgets first
        _load_qt_widgets()

        # Initialize user management (lightweight)
        self.user_management = UserManagement()

        # Create menu bar
        self.create_menus()

        # Create stacked widget for multiple screens
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Create and add login screen
        self.show_login_on_start = False  # Set to False to bypass login at startup

        # Defer heavy UI component creation until needed
        self._main_dashboard = None
        self._split_screen_processor = None
        self._template_manager = None
        self._bulk_processor = None

        # Initialize the main dashboard (this is lightweight)
        self._create_main_dashboard()

        # Initialize other objects as None first, then create them lazily
        self.pdf_processor = None
        self.template_manager = None
        self.multi_page_processor = None
        self.invoice_viewer = None

        # Set window to maximized state by default
        self.showMaximized()
        print(f"[DEBUG] Setting default window state to maximized")

        # Initialize window state
        self.is_fullscreen = self.isFullScreen()
        self.is_maximized = self.isMaximized()
        print(f"[DEBUG] Initial window state - fullscreen: {self.is_fullscreen}, maximized: {self.is_maximized}")

        # Initialize menu text based on window state
        self.update_window_state_menu_text()

        # Add keyboard shortcuts for window state toggles
        from PySide6.QtGui import QShortcut, QKeySequence

        # F11 for fullscreen toggle
        self.fullscreen_shortcut = QShortcut(QKeySequence("F11"), self)
        self.fullscreen_shortcut.activated.connect(self.toggle_fullscreen)
        print(f"[DEBUG] Added F11 shortcut for fullscreen toggle")

        # F10 for maximize toggle
        self.maximize_shortcut = QShortcut(QKeySequence("F10"), self)
        self.maximize_shortcut.activated.connect(self.toggle_maximized)
        print(f"[DEBUG] Added F10 shortcut for maximize toggle")

        # Check if we should show login on start
        if self.show_login_on_start:
            self.handle_login_request()
        else:
            self.stacked_widget.setCurrentWidget(self.main_dashboard)

    def _create_main_dashboard(self):
        """Create the main dashboard lazily"""
        if self._main_dashboard is None:
            self._main_dashboard = MainDashboard(self.user_management)
            self._main_dashboard.show_pdf_processor.connect(self.show_pdf_processor)
            self._main_dashboard.show_template_manager.connect(self.show_template_manager)
            self._main_dashboard.show_bulk_processor.connect(self.show_bulk_processor)
            self._main_dashboard.show_user_management.connect(self.show_user_management)
            self._main_dashboard.user_profile.login_requested.connect(self.handle_login_request)
            self._main_dashboard.user_profile.logout_requested.connect(self.update_menus)
            self._main_dashboard.login_successful.connect(self.handle_login_success)
            self.stacked_widget.addWidget(self._main_dashboard)
        return self._main_dashboard

    @property
    def main_dashboard(self):
        """Get the main dashboard, creating it if needed"""
        return self._create_main_dashboard()

    def _create_split_screen_processor(self):
        """Create the split screen processor lazily"""
        if self._split_screen_processor is None:
            self._split_screen_processor = SplitScreenInvoiceProcessor()
            self.stacked_widget.addWidget(self._split_screen_processor)

            # Connect signals
            try:
                self._split_screen_processor.go_back.connect(self.handle_split_screen_processor_go_back)
                print(f"[DEBUG] Successfully connected split_screen_processor go_back signal")
            except Exception as e:
                print(f"[DEBUG] Error connecting split_screen_processor go_back signal: {str(e)}")

            try:
                self._split_screen_processor.save_template_signal.connect(self.handle_template_saved)
                print(f"[DEBUG] Successfully connected split_screen_processor save_template_signal")
            except Exception as e:
                print(f"[DEBUG] Error connecting split_screen_processor save_template_signal: {str(e)}")

            # Set as PDF processor
            self.pdf_processor = self._split_screen_processor

        return self._split_screen_processor

    @property
    def split_screen_processor(self):
        """Get the split screen processor, creating it if needed"""
        return self._create_split_screen_processor()

    def _create_template_manager(self):
        """Create the template manager lazily"""
        if self._template_manager is None:
            # Ensure we have a PDF processor first
            pdf_proc = self.split_screen_processor  # This will create it if needed
            self._template_manager = TemplateManager(pdf_proc)
            self.stacked_widget.addWidget(self._template_manager)

            # Connect signals
            self._template_manager.go_back.connect(self.handle_template_manager_go_back)
            self._template_manager.template_selected.connect(self.apply_template)

            # Set as template manager
            self.template_manager = self._template_manager

        return self._template_manager

    @property
    def bulk_processor(self):
        """Get the bulk processor, creating it if needed"""
        if self._bulk_processor is None:
            self._bulk_processor = BulkProcessor()
            self._bulk_processor.go_back.connect(self.handle_bulk_processor_go_back)
            self.stacked_widget.addWidget(self._bulk_processor)
        return self._bulk_processor

    def create_menus(self):
        """Create application menu bar with various options."""
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)

        # File menu
        file_menu = menubar.addMenu('&File')

        # Exit action
        exit_action = QAction('E&xit', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.setStatusTip('Exit application')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Admin menu (only for developers with user_management permission)
        self.admin_menu = menubar.addMenu('&Admin')

        # User management action
        self.user_management_action = QAction('&User Management', self)
        self.user_management_action.setStatusTip('Manage users and permissions')
        self.user_management_action.triggered.connect(self.show_user_management)
        self.admin_menu.addAction(self.user_management_action)

        # Role management action
        self.role_management_action = QAction('&Role Management', self)
        self.role_management_action.setStatusTip('Manage roles and their permissions')
        self.role_management_action.triggered.connect(self.show_role_management)
        self.admin_menu.addAction(self.role_management_action)

        # Add separator
        self.admin_menu.addSeparator()

        # License management action
        self.license_management_action = QAction('&License Management', self)
        self.license_management_action.setStatusTip('Manage software licensing')
        self.license_management_action.triggered.connect(self.show_license_management)
        self.admin_menu.addAction(self.license_management_action)

        # Initially disable admin menu
        self.admin_menu.setEnabled(False)

        # View menu
        view_menu = menubar.addMenu('&View')

        # Fullscreen action
        fullscreen_action = QAction('&Fullscreen', self)
        fullscreen_action.setShortcut('F11')
        fullscreen_action.setStatusTip('Toggle fullscreen mode')
        fullscreen_action.triggered.connect(self.toggle_fullscreen)
        view_menu.addAction(fullscreen_action)

        # Maximize action
        self.maximize_action = QAction('&Maximize', self)
        self.maximize_action.setShortcut('F10')
        self.maximize_action.setStatusTip('Toggle maximized window')
        self.maximize_action.triggered.connect(self.toggle_maximized)
        view_menu.addAction(self.maximize_action)

        # Help menu
        help_menu = menubar.addMenu('&Help')

        # About action
        about_action = QAction('&About', self)
        about_action.setStatusTip('About this application')
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def update_menus(self):
        """Update menu availability based on user permissions."""
        if self.user_management.get_current_user() and self.user_management.has_permission('user_management'):
            self.admin_menu.setEnabled(True)
        else:
            self.admin_menu.setEnabled(False)

    def handle_login_request(self):
        """Show the login panel in the dashboard."""
        # Switch to main dashboard if not already there
        self.stacked_widget.setCurrentWidget(self.main_dashboard)
        # Update dashboard to show login panel
        self.main_dashboard.update_dashboard()
        # Focus on username field
        if hasattr(self.main_dashboard, 'username_edit'):
            self.main_dashboard.username_edit.setFocus()

    def handle_login_success(self, user):
        """Handle successful login."""
        # Reset the state of all screens first
        for i in range(self.stacked_widget.count()):
            widget = self.stacked_widget.widget(i)
            if widget != self.main_dashboard:
                widget.setVisible(False)

        # Make sure the main dashboard is visible and in front
        self.main_dashboard.setVisible(True)
        self.stacked_widget.setCurrentWidget(self.main_dashboard)

        # Update the main dashboard to reflect the user's permissions
        self.main_dashboard.update_dashboard()

        # Update menus based on user permissions
        self.update_menus()

        # Show welcome message - commented out as requested
        # QMessageBox.information(
        #     self,
        #     "Login Successful",
        #     f"Welcome, {user['full_name']}!\n\nYou are logged in as: {user['role_name']}"
        # )

    def show_pdf_processor(self):
        """Show the PDF processor screen with proper transition."""
        # Check if user has permission
        if (not self.user_management.get_current_user() or
            not self.user_management.has_permission("draw_pdf_rules")):
            result = QMessageBox.question(
                self,
                "Permission Required",
                "You need to log in with appropriate permissions to access this feature.\n\nWould you like to log in now?",
                QMessageBox.Yes | QMessageBox.No
            )

            if result == QMessageBox.Yes:
                # Show login panel in dashboard
                self.handle_login_request()
            return

        # Use unified screen switching method
        self._switch_to_screen(self.split_screen_processor)

    def _switch_to_screen(self, target_widget):
        """Unified method to switch to any screen with proper state management."""
        # Update window state before switching
        self.update_window_state()

        # Hide all widgets in the stacked widget to prevent overlapping
        for i in range(self.stacked_widget.count()):
            widget = self.stacked_widget.widget(i)
            if widget:
                widget.setVisible(False)

        # Show the target widget
        target_widget.setVisible(True)
        self.stacked_widget.setCurrentWidget(target_widget)

        # Apply the stored window state
        self.apply_window_state()

    def show_template_manager(self):
        """Show the template manager screen."""
        # Only allow if user has permission
        if (not self.user_management.get_current_user() or
            not self.user_management.has_permission("template_management")):
            UIMessageFactory.show_warning(
                self,
                "Permission Denied",
                "You do not have permission to access template management.\n\n"
                "Please log in with a developer account to access this feature."
            )
            return

        # Use unified screen switching method
        self._switch_to_screen(self._create_template_manager())

    def show_bulk_processor(self):
        """Show the bulk processor screen."""
        # Check if user has permission
        if (not self.user_management.get_current_user() or
            not self.user_management.has_permission("bulk_extraction")):
            UIMessageFactory.show_warning(
                self,
                "Permission Denied",
                "You do not have permission to access bulk extraction.\n\n"
                "Please log in with an appropriate account to access this feature."
            )
            return

        # Use unified screen switching method
        self._switch_to_screen(self.bulk_processor)

    def _handle_go_back_to_dashboard(self, source_name: str = ""):
        """Unified method to handle go_back signals from any processor."""
        if source_name:
            print(f"[DEBUG] Handling {source_name} go_back signal")

        # Update window state before switching
        self.update_window_state()

        # Hide all widgets first
        for i in range(self.stacked_widget.count()):
            widget = self.stacked_widget.widget(i)
            if widget:
                widget.setVisible(False)

        # Show dashboard and make it current
        self.main_dashboard.setVisible(True)
        self.stacked_widget.setCurrentWidget(self.main_dashboard)

        # Apply the stored window state
        self.apply_window_state()

    def handle_bulk_processor_go_back(self):
        """Handle the go_back signal from the bulk processor."""
        self._handle_go_back_to_dashboard("bulk_processor")

    def handle_template_manager_go_back(self):
        """Handle the go_back signal from the template manager."""
        self._handle_go_back_to_dashboard("template_manager")

    def handle_split_screen_processor_go_back(self):
        """Handle the go_back signal from the split screen processor."""
        self._handle_go_back_to_dashboard("split_screen_processor")

    def handle_template_saved(self):
        """Handle template saved signal from split screen processor."""
        print(f"[DEBUG] Template saved signal received - refreshing template manager")

        # Refresh the template manager to show the newly saved template
        if self._template_manager is not None:
            try:
                self._template_manager.refresh()
                print(f"[DEBUG] Template manager refreshed successfully")
            except Exception as e:
                print(f"[ERROR] Failed to refresh template manager: {e}")
        else:
            print(f"[WARNING] Template manager not available for refresh")

    def apply_template(self, template):
        """Apply the selected template to the current PDF processor"""
        try:
            print("\n[DEBUG] apply_template called in main.py")
            print(f"[DEBUG] Template name: {template.get('name')}")
            print(f"[DEBUG] Template type: {template.get('template_type', 'single')}")
            print(f"[DEBUG] Current widget: {self.stacked_widget.currentWidget().__class__.__name__}")

            # Get the PDF file path from the template data
            file_path = template.get('selected_pdf_path')

            if not file_path:
                # If no file path in template data, ask user to select a file
                file_path, _ = QFileDialog.getOpenFileName(
                    self, "Open PDF File", "", "PDF Files (*.pdf)"
                )
                if not file_path:
                    # User cancelled, don't apply template
                    return

            # Load the PDF file with skip_dialog=True to avoid the multi-page dialog
            # Set source="template_manager" to indicate this PDF is being loaded from a template
            self.pdf_processor.load_pdf_file(file_path, skip_dialog=True, source="template_manager")
            print(f"Loaded PDF file: {file_path}")

            # Ensure the PDF processor is visible and has focus
            self.stacked_widget.setCurrentWidget(self.pdf_processor)

            # Apply template directly without complex page mapping logic
            print(f"[DEBUG] Applying template with simplified page-wise logic")

            # Set template type and basic configuration
            if template.get('template_type') == 'multi':
                self.pdf_processor.multi_page_mode = True
                self.pdf_processor.prev_page_btn.show()
                self.pdf_processor.next_page_btn.show()
                self.pdf_processor.apply_to_remaining_btn.show()
            else:
                self.pdf_processor.multi_page_mode = False
                self.pdf_processor.prev_page_btn.hide()
                self.pdf_processor.next_page_btn.hide()
                self.pdf_processor.apply_to_remaining_btn.hide()

            # Apply regions and column lines directly
            if template.get('template_type') == 'single':
                # For single-page templates, apply to current page
                print(f"[DEBUG] Applying single-page template")
                self._apply_single_page_template_direct(template)
            else:
                # For multi-page templates, apply page by page
                print(f"[DEBUG] Applying multi-page template")
                self._apply_multi_page_template_direct(template)

            # Set extraction parameters
            if 'config' in template:
                print(f"[DEBUG] Setting extraction parameters from config")
                self.pdf_processor.extraction_params = template['config'].copy()
                self.pdf_processor._ensure_extraction_params_structure()

            # Force update display
            print(f"[DEBUG] Forcing PDF display update after template application")
            self.pdf_processor.pdf_label.update()
            self.pdf_processor.pdf_label.repaint()

            # Update extraction results
            self.pdf_processor.update_extraction_results(force=True)

            # Debug: Check if regions are properly set
            print(f"[DEBUG] Regions after template application:")
            for section, region_list in self.pdf_processor.regions.items():
                print(f"  {section}: {len(region_list)} regions")
                for i, region in enumerate(region_list):
                    if hasattr(region, 'rect') and hasattr(region, 'label'):
                        print(f"    {i}: {region.label} - {region.rect}")
                    else:
                        print(f"    {i}: {type(region)} - {region}")

            # Show success message
            success_msg = QMessageBox(self)
            success_msg.setWindowTitle("Template Applied")
            success_msg.setText("Template Applied Successfully")
            success_msg.setInformativeText(f"The template '{template['name']}' has been applied to '{os.path.basename(file_path)}'.")
            success_msg.setIcon(QMessageBox.Information)
            success_msg.setStyleSheet("QLabel { color: black; }")
            success_msg.exec()

            # Switch back to the PDF processor screen
            self.stacked_widget.setCurrentWidget(self.pdf_processor)

        except Exception as e:
            print(f"[ERROR] Error applying template: {str(e)}")
            import traceback
            traceback.print_exc()

            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Error Applying Template")
            error_dialog.setInformativeText(f"An error occurred: {str(e)}")
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setStyleSheet("QLabel { color: black; }")
            error_dialog.exec()

    def _apply_single_page_template_direct(self, template):
        """Apply single-page template directly to the PDF processor"""
        try:
            from standardized_coordinates import StandardRegion
            from PySide6.QtCore import QRect

            # Convert template regions to StandardRegion objects
            if 'regions' in template:
                print(f"[DEBUG] Converting template regions to StandardRegion objects")
                print(f"[DEBUG] Template regions data: {template['regions']}")
                self.pdf_processor.regions = {'header': [], 'items': [], 'summary': []}

                for section, region_list in template['regions'].items():
                    print(f"[DEBUG] Processing {section} section with {len(region_list) if region_list else 0} regions")
                    if not region_list:
                        continue

                    for i, region in enumerate(region_list):
                        print(f"[DEBUG] Processing region {i} in {section}: {region}")
                        print(f"[DEBUG] Region type: {type(region)}")

                        if isinstance(region, StandardRegion):
                            # Already a StandardRegion, use directly
                            self.pdf_processor.regions[section].append(region)
                            print(f"[DEBUG] Using existing StandardRegion {region.label}")
                        elif hasattr(region, 'rect') and hasattr(region, 'label'):
                            # StandardRegion-like object, use directly
                            self.pdf_processor.regions[section].append(region)
                            print(f"[DEBUG] Using StandardRegion-like object {region.label}")
                        elif isinstance(region, QRect):
                            # QRect object, convert to StandardRegion
                            label = f"{section[0].upper()}{i+1}"
                            print(f"[DEBUG] QRect coordinates: x={region.x()}, y={region.y()}, w={region.width()}, h={region.height()}")
                            if region.width() <= 0 or region.height() <= 0:
                                print(f"[ERROR] Invalid QRect dimensions: {region.width()}x{region.height()}, skipping region")
                                continue
                            standard_region = self.pdf_processor.create_standard_region(
                                region.x(), region.y(), region.width(), region.height(), section, i
                            )
                            self.pdf_processor.regions[section].append(standard_region)
                            print(f"[DEBUG] Converted QRect to StandardRegion {label}")
                        elif isinstance(region, dict) and all(k in region for k in ['x', 'y', 'width', 'height']):
                            # Dictionary format, convert to StandardRegion
                            label = region.get('label', f"{section[0].upper()}{i+1}")
                            x, y, w, h = int(region['x']), int(region['y']), int(region['width']), int(region['height'])
                            print(f"[DEBUG] Dict coordinates: x={x}, y={y}, w={w}, h={h}")
                            if w <= 0 or h <= 0:
                                print(f"[ERROR] Invalid dict dimensions: {w}x{h}, skipping region")
                                continue
                            standard_region = self.pdf_processor.create_standard_region(x, y, w, h, section, i)
                            self.pdf_processor.regions[section].append(standard_region)
                            print(f"[DEBUG] Converted dict to StandardRegion {label}")
                        else:
                            print(f"[WARNING] Unknown region format: {type(region)}, content: {region}")

                # Update region labels for consistency
                for section in self.pdf_processor.regions:
                    self.pdf_processor._update_region_labels(section)

            # Apply column lines
            if 'column_lines' in template:
                print(f"[DEBUG] Applying column lines from template")
                self.pdf_processor.column_lines = template['column_lines'].copy()

        except Exception as e:
            print(f"[ERROR] Error applying single-page template: {str(e)}")
            import traceback
            traceback.print_exc()

    def _apply_multi_page_template_direct(self, template):
        """Apply multi-page template directly to the PDF processor"""
        try:
            print(f"[DEBUG] Applying multi-page template (simplified)")

            # For now, apply the first page configuration to all pages
            # This will be enhanced with proper page mapping later
            if 'page_regions' in template and template['page_regions']:
                first_page_regions = template['page_regions'][0]
                # Apply first page regions as single-page template
                single_page_template = {
                    'regions': first_page_regions,
                    'column_lines': template.get('page_column_lines', [{}])[0] if template.get('page_column_lines') else {},
                    'config': template.get('config', {})
                }
                self._apply_single_page_template_direct(single_page_template)
            elif 'regions' in template:
                # Fallback to single-page logic
                self._apply_single_page_template_direct(template)

        except Exception as e:
            print(f"[ERROR] Error applying multi-page template: {str(e)}")
            import traceback
            traceback.print_exc()

    def show_template_manager_from_viewer(self):
        """Show the template manager from the invoice viewer, preserving all data for template creation"""
        print("\n[DEBUG] show_template_manager_from_viewer called")

        # Check if we have a valid template manager
        if not hasattr(self, 'template_manager') or self.template_manager is None:
            print("[ERROR] Template manager is not available")
            QMessageBox.critical(
                self,
                "Error",
                "Template manager is not available. Please try again."
            )
            return

        # Update window state before switching
        self.update_window_state()

        # Make sure the current invoice configuration is retained
        # This ensures that when we save a template, we're saving the current state
        try:
            self.template_manager.refresh()

            # Hide all widgets in the stacked widget to prevent overlapping
            for i in range(self.stacked_widget.count()):
                widget = self.stacked_widget.widget(i)
                if widget:
                    widget.setVisible(False)

            # Make sure the template manager is visible
            self.template_manager.setVisible(True)

            # Switch to the template manager
            self.stacked_widget.setCurrentWidget(self.template_manager)

            # Apply the stored window state
            self.apply_window_state()

            # Show a hint to the user about template saving
            QMessageBox.information(
                self,
                "Save Template",
                "You can now save the current invoice configuration as a template.\n\n"
                "Templates save all table regions, column lines, and extraction settings "
                "for reuse with similar invoices."
            )
        except Exception as e:
            print(f"[ERROR] Failed to switch to template manager: {str(e)}")
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to switch to template manager: {str(e)}"
            )

    def show_user_management(self):
        """Show the user management dialog."""
        # Check if user has permission
        if (not self.user_management.get_current_user() or
            not self.user_management.has_permission("user_management")):
            QMessageBox.warning(
                self,
                "Permission Denied",
                "You do not have permission to access user management.\n\n"
                "Please log in with an administrator account to access this feature."
            )
            return

        dialog = UserManagementDialog(self.user_management, self)
        dialog.exec()

    def show_role_management(self):
        """Show the role management dialog."""
        # Check if user has permission
        if (not self.user_management.get_current_user() or
            not self.user_management.has_permission("user_management")):
            QMessageBox.warning(
                self,
                "Permission Denied",
                "You do not have permission to access role management.\n\n"
                "Please log in with an administrator account to access this feature."
            )
            return

        dialog = RoleManagementDialog(self.user_management, self)
        dialog.exec()

    def show_license_management(self):
        """Show the license management dialog."""
        # Check if user has permission
        if (not self.user_management.get_current_user() or
            not self.user_management.has_permission("user_management")):
            QMessageBox.warning(
                self,
                "Permission Denied",
                "You do not have permission to access license management.\n\n"
                "Please log in with an administrator account to access this feature."
            )
            return

        # Show the license management dialog with admin privileges
        show_license_manager_dialog(is_admin=True)

    def show_about_dialog(self):
        """Show the about dialog."""
        QMessageBox.about(
            self,
            "About PDF Harvest",
            "PDF Harvest\n\n"
            "Version 1.0.0\n\n"
            "A PDF invoice data extraction tool with visual selection and mapping capabilities.\n\n"
            "© 2023 PDF Harvest Team"
            )

    # Rules manager methods removed as they are no longer needed

    def update_window_state(self):
        """Update the stored window state based on current window state"""
        # Check fullscreen state
        current_fullscreen = self.isFullScreen()
        if current_fullscreen != self.is_fullscreen:
            self.is_fullscreen = current_fullscreen
            print(f"[DEBUG] Window state updated - fullscreen: {self.is_fullscreen}")

        # Check maximized state (only if not in fullscreen)
        if not current_fullscreen:
            current_maximized = self.isMaximized()
            if current_maximized != self.is_maximized:
                self.is_maximized = current_maximized
                print(f"[DEBUG] Window state updated - maximized: {self.is_maximized}")

    def apply_window_state(self):
        """Apply the stored window state to the current window"""
        if self.is_fullscreen and not self.isFullScreen():
            print(f"[DEBUG] Applying fullscreen state")
            self.showFullScreen()
        elif not self.is_fullscreen and self.isFullScreen():
            print(f"[DEBUG] Applying normal window state")
            self.showNormal()
            # Apply maximized state if needed
            if self.is_maximized and not self.isMaximized():
                print(f"[DEBUG] Applying maximized state")
                self.showMaximized()

    def toggle_fullscreen(self):
        """Toggle between fullscreen and normal window state"""
        if self.isFullScreen():
            # Going from fullscreen to normal
            self.is_fullscreen = False
            self.showNormal()
            print(f"[DEBUG] Toggled to normal window state")
            # Restore maximized state if it was maximized before
            if self.is_maximized:
                self.showMaximized()
                print(f"[DEBUG] Restored maximized state")
        else:
            # Going from normal/maximized to fullscreen
            # Remember if we were maximized
            self.is_maximized = self.isMaximized()
            self.is_fullscreen = True
            self.showFullScreen()
            print(f"[DEBUG] Toggled to fullscreen state (was maximized: {self.is_maximized})")

    def toggle_maximized(self):
        """Toggle between maximized and normal window state"""
        if self.isFullScreen():
            # If in fullscreen, exit fullscreen first
            self.is_fullscreen = False
            self.showNormal()
            self.is_maximized = True
            self.showMaximized()
            print(f"[DEBUG] Exited fullscreen and maximized window")
        elif self.isMaximized():
            # Going from maximized to normal
            self.is_maximized = False
            self.showNormal()
            print(f"[DEBUG] Toggled to normal window state")
        else:
            # Going from normal to maximized
            self.is_maximized = True
            self.showMaximized()
            print(f"[DEBUG] Toggled to maximized state")

    def changeEvent(self, event):
        """Handle window state change events to track fullscreen state"""
        # In PySide6, we need to use QEvent.Type.WindowStateChange
        from PySide6.QtCore import QEvent
        if event.type() == QEvent.Type.WindowStateChange:
            self.update_window_state()
            # Update menu text based on current state
            self.update_window_state_menu_text()
        super().changeEvent(event)

    def update_window_state_menu_text(self):
        """Update menu text to reflect current window state"""
        if hasattr(self, 'maximize_action'):
            if self.isMaximized():
                self.maximize_action.setText('&Restore')
                self.maximize_action.setStatusTip('Restore window to normal size')
            else:
                self.maximize_action.setText('&Maximize')
                self.maximize_action.setStatusTip('Maximize window')

    def closeEvent(self, event):
        """Handle the window close event to ensure proper cleanup."""
        print("\nApplication window closing, performing cleanup...")

        # Close all database connections
        self.close_all_database_connections()

        # Accept the close event
        event.accept()

    def close_all_database_connections(self):
        """Close all database connections in the application."""
        # Close user management database connection
        if hasattr(self, 'user_management') and self.user_management:
            try:
                print("Closing user management database connection...")
                self.user_management.close()
                print("User management database connection closed successfully")
            except Exception as e:
                print(f"Warning: Error closing user management connection: {str(e)}")

        # Close template manager database connection
        if hasattr(self, 'template_manager') and self.template_manager:
            try:
                if hasattr(self.template_manager, 'db') and self.template_manager.db:
                    print("Closing template manager database connection...")
                    self.template_manager.db.close()
                    print("Template manager database connection closed successfully")
            except Exception as e:
                print(f"Warning: Error closing template manager connection: {str(e)}")

# Create logs directory if it doesn't exist
def ensure_logs_directory():
    """Create logs directory if it doesn't exist"""
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    if not os.path.exists(logs_dir):
        try:
            os.makedirs(logs_dir)
            print(f"Created logs directory: {logs_dir}")
        except Exception as e:
            print(f"Warning: Could not create logs directory: {str(e)}")
    return logs_dir

# Global exception handler
def global_exception_handler(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions by logging them and showing a message to the user"""
    # Get the current timestamp for the log filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Ensure logs directory exists
    logs_dir = ensure_logs_directory()
    log_file = os.path.join(logs_dir, f"error_{timestamp}.log")

    # Format the exception information
    exception_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))

    # Log the exception to a file
    try:
        with open(log_file, "w") as f:
            f.write(f"Exception occurred at {timestamp}\n\n")
            f.write(exception_text)
            f.write("\n\nSystem Information:\n")
            f.write(f"Python version: {sys.version}\n")
            f.write(f"Operating system: {os.name} - {sys.platform}\n")

        print(f"Exception logged to: {log_file}")
    except Exception as e:
        print(f"Error writing to log file: {str(e)}")
        print(exception_text)  # Print to console as fallback

    # Show a message box to the user
    try:
        # Only show the message box if QApplication exists
        if QApplication.instance():
            error_msg = f"An unexpected error occurred:\n\n{str(exc_value)}\n\nThe error has been logged to:\n{log_file}\n\nPlease report this issue to the developer."
            QMessageBox.critical(None, "Application Error", error_msg)
    except Exception as msg_error:
        print(f"Error showing message box: {str(msg_error)}")

    # Call the original exception handler
    sys.__excepthook__(exc_type, exc_value, exc_traceback)

# Global variable to store the main window reference
main_window = None

# Function to handle application cleanup
def cleanup_application():
    """Perform cleanup tasks before application exit"""
    print("\nPerforming application cleanup...")

    global main_window

    # Close all database connections if window exists
    if main_window is not None:
        try:
            # Use the window's method to close all database connections
            main_window.close_all_database_connections()
        except Exception as e:
            print(f"Warning: Error during database connection cleanup: {str(e)}")
    else:
        print("No main window reference found during cleanup")

    # Now run the database protection cleanup
    try:
        cleanup_database_protection()
        print("Database protection cleanup completed successfully")
    except Exception as e:
        print(f"Warning: Error during database protection cleanup: {str(e)}")
        import traceback
        traceback.print_exc()

def run_cli_command(args):
    """Run a command-line interface command"""
    try:
        # Parse CLI arguments
        parser = argparse.ArgumentParser(description='Bulk PDF data extraction using templates')
        parser.add_argument('--folder', required=True, help='Folder containing PDF files to process')
        parser.add_argument('--template', required=True, help='Template name to use for extraction')
        parser.add_argument('--username', required=True, help='Username for authentication')
        parser.add_argument('--password', required=True, help='Password for authentication')
        parser.add_argument('--output', help='Output directory for extracted data (optional)')
        parser.add_argument('--threads', type=int, help='Number of threads to use for parallel processing (default: CPU count)')
        parser.add_argument('--chunk', type=int, help='Chunk size for large documents (default: 50)')

        cli_args = parser.parse_args(args)

        # Run the CLI command
        result = process_pdf_folder(
            cli_args.folder,
            cli_args.template,
            cli_args.username,
            cli_args.password,
            cli_args.output,
            cli_args.threads,
            cli_args.chunk
        )

        # Return success/failure code
        return 0 if result else 1
    except Exception as e:
        print(f"CLI command error: {str(e)}")
        traceback.print_exc()
        return 1

def setup_logging():
    """Set up logging to file"""
    import logging
    import time

    # Create logs directory if it doesn't exist
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)

    # Set up logging
    log_file = os.path.join(logs_dir, f"app_{time.strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # Add console handler
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(name)s - %(levelname)s - %(message)s')
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)

    return logging.getLogger('main')

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description="PDFHarvest - PDF Data Extraction Tool")
    parser.add_argument("--debug", action="store_true", help="Enable debug mode with verbose logging")
    parser.add_argument("--cli", action="store_true", help="Run in command-line interface mode")
    parser.add_argument("--version", action="store_true", help="Show version information and exit")

    # Add any additional CLI arguments here

    # Parse known args only, to avoid errors with Qt's own arguments
    args, unknown = parser.parse_known_args()
    return args

def main():
    """Main application entry point"""
    # Parse command line arguments
    args = parse_arguments()

    # Set up logging
    logger = setup_logging()
    logger.info("Application starting")
    logger.info(f"Python version: {sys.version}")
    logger.info(f"Current directory: {os.getcwd()}")
    logger.info(f"Command line arguments: {sys.argv}")

    # Show version if requested
    if args.version:
        print("PDFHarvest version 1.0.0")
        return 0

    try:
        # Install the global exception handler
        sys.excepthook = global_exception_handler

        # Ensure logs directory exists
        ensure_logs_directory()

        print("Starting PDFHarvest application...")
        logger.info("Starting PDFHarvest application...")

        # Check for CLI mode
        if args.cli:
            # Get all arguments after --cli
            cli_index = sys.argv.index("--cli") if "--cli" in sys.argv else -1
            cli_args = sys.argv[cli_index+1:] if cli_index >= 0 else []
            logger.info(f"Running in CLI mode with args: {cli_args}")

            # Initialize database protection
            if not initialize_database_protection():
                logger.error("Failed to initialize the user database")
                print("Failed to initialize the user database. The application cannot start.")
                return 1

            # Run in CLI mode
            exit_code = run_cli_command(cli_args)

            # Clean up
            cleanup_application()

            # Exit with the appropriate code
            return exit_code

        # Check for debug mode
        debug_mode = args.debug
        if debug_mode:
            # Set root logger to DEBUG level
            logging.getLogger().setLevel(logging.DEBUG)
            # Set console handler to DEBUG level for all loggers
            for handler in logging.getLogger().handlers:
                if isinstance(handler, logging.StreamHandler):
                    handler.setLevel(logging.DEBUG)

            logger.info("Debug mode enabled - verbose logging activated")
            logger.debug("Debug messages will now be displayed")

        logger.info("Creating QApplication")
        app = QApplication(sys.argv)
        app.setStyle('Fusion')

        # Register cleanup function to run on application exit
        logger.info("Registering cleanup handler")
        app.aboutToQuit.connect(cleanup_application)

        # Defer heavy operations until after QApplication is created
        # This allows the app to start faster and show a loading screen if needed

        # Load heavy modules now that QApplication exists
        logger.info("Loading heavy modules")
        _load_heavy_modules()

        # Initialize database protection
        logger.info("Initializing database protection")
        if not initialize_database_protection():
            logger.error("Failed to initialize the user database")
            QMessageBox.critical(None, "Database Error",
                              "Failed to initialize the user database.\n\n"
                              "The application cannot start.")
            return 1

        # Check if the application is activated
        # For initial activation, we don't have a user yet, so is_admin=False
        logger.info("Checking activation")
        if check_activation(is_admin=False) or debug_mode:
            logger.info("Activation successful or debug mode enabled")

            # Create main window with additional debug logging
            logger.debug("Creating main application window")
            global main_window
            try:
                main_window = PDFHarvest()
                logger.debug("Main window created successfully")

                # Show the main window
                logger.debug("Showing main window")
                main_window.show()
                logger.debug("Main window displayed")

                # Enter the main event loop
                logger.info("Entering main event loop")
                return app.exec()
            except Exception as e:
                logger.critical(f"Failed to create or show main window: {str(e)}")
                logger.critical(traceback.format_exc())
                raise
        else:
            # Exit if not activated
            logger.error("Activation required")
            QMessageBox.critical(None, "Activation Required",
                              "This software requires activation to run.\n\n"
                              "Please contact your vendor to obtain a license key.")
            return 1

    except Exception as e:
        # This is a fallback in case the global exception handler fails
        if 'logger' in locals():
            logger.critical(f"Critical error during application startup: {str(e)}")
            logger.critical(traceback.format_exc())

        error_message = f"Critical error during application startup: {str(e)}\n\n"
        error_message += "Please check the logs directory for details."

        # Try to show a message box if possible
        try:
            if QApplication.instance():
                QMessageBox.critical(None, "Critical Error", error_message)
            else:
                # If QApplication doesn't exist yet, create a temporary one
                temp_app = QApplication(sys.argv)
                QMessageBox.critical(None, "Critical Error", error_message)
        except:
            # If all else fails, print to console
            print("CRITICAL ERROR:", error_message)
            traceback.print_exc()

        return 1

if __name__ == '__main__':
    sys.exit(main())

