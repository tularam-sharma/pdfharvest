from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QFileDialog, QScrollArea, QSplitter, QMessageBox,
                             QFrame, QStackedWidget, QTreeWidget, QTreeWidgetItem, QDialog,
                             QFormLayout, QSpinBox, QCheckBox, QLineEdit, QDialogButtonBox,
                             QGroupBox, QComboBox, QDoubleSpinBox, QInputDialog, QTextEdit,
                             QTabWidget, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
                             QHeaderView, QApplication, QMenu, QToolTip)
from PySide6.QtCore import Qt, Signal, QRect, QPoint, QSize, QEvent, QRegularExpression
from PySide6.QtGui import (QPixmap, QPainter, QPen, QColor, QCursor, QFont, QImage,
                          QKeySequence, QShortcut, QTextCharFormat, QTextCursor, QBrush)
import fitz  # PyMuPDF
from PIL import Image
import io
import os
import sys
import pandas as pd
import numpy as np
import json
import enum
import copy
import re
import yaml
import subprocess
import datetime
import decimal
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Any, Union

# Import factory modules for code deduplication
from common_factories import (
    TemplateFactory, DatabaseOperationFactory, UIMessageFactory,
    ValidationFactory, get_database_factory
)
from ui_component_factory import UIComponentFactory, LayoutFactory
from simplified_extraction_engine import get_extraction_engine

# Import path_helper for path resolution (optional)
try:
    import path_helper
    PATH_HELPER_AVAILABLE = True
except ImportError:
    PATH_HELPER_AVAILABLE = False
    print("[WARNING] path_helper not available. Using default path resolution.")

# Custom JSON encoder to handle non-serializable objects
class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime.datetime, datetime.date)):
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

# Import simplified utilities
from pdf_extraction_utils import (extract_table, extract_tables, clean_dataframe, DEFAULT_EXTRACTION_PARAMS,
                                convert_display_to_pdf_coords, convert_pdf_to_display_coords, get_scale_factors,
                                clear_extraction_cache, clear_extraction_cache_for_pdf, clear_extraction_cache_for_section,
                                get_extraction_cache_stats)
from multi_method_extraction import extract_with_method, cleanup_extraction
import pypdf_table_extraction
# Import simplified invoice processing utilities
import invoice_processing_utils as invoice_processing_utils

# Import unified extraction components
try:
    from extraction_adapters import extract_from_user_regions, process_invoice2data_unified
    from common_extraction_engine import get_extraction_engine
    from region_adapters import create_region_adapter
    UNIFIED_EXTRACTION_AVAILABLE = True
except ImportError as e:
    print(f"[WARNING] Unified extraction not available: {e}")
    UNIFIED_EXTRACTION_AVAILABLE = False

# Import cache manager
try:
    from cache_manager import get_cache_manager, register_cleanup_callback
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

# Import standardized coordinate system - NO backward compatibility
from standardized_coordinates import StandardRegion
from coordinate_boundary_converters import DatabaseBoundaryConverter
from region_utils import validate_rect
from single_format_region_utils import extract_rect_and_label, create_region_item
from error_handler import log_error, log_warning, handle_exception, ErrorContext
from extraction_params_utils import (
    normalize_extraction_params, prepare_section_params, create_standardized_extraction_call,
    ExtractionParamsHandler
)
from region_label_utils import (
    RegionLabelHandler, create_region_label, standardize_dataframe_labels, get_display_label
)

# Import PDF extraction utilities
from pdf_extraction_utils import (
    extract_table, extract_tables, clean_dataframe, get_scale_factors,
    clear_extraction_cache_for_section, clear_extraction_cache_for_pdf, clear_extraction_cache
)

# Import invoice processing utilities
try:
    import invoice_processing_utils
    INVOICE_PROCESSING_UTILS_AVAILABLE = True
except ImportError:
    INVOICE_PROCESSING_UTILS_AVAILABLE = False
    print("Warning: invoice_processing_utils not available. Some template mapping features may be limited.")

# path_helper already imported above

# Define RegionType enum to replace the import from pdf_processor
class RegionType(enum.Enum):
    HEADER = "header"
    ITEMS = "items"
    SUMMARY = "summary"


class ClickableLabel(QLabel):
    """A QLabel that emits a clicked signal when clicked"""
    clicked = Signal()

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


class PDFLabel(QLabel):
    """Custom QLabel for displaying PDFs with drawing capabilities"""
    # Define the signal at the class level
    zoom_changed = Signal(float)  # Signal to notify when zoom changes

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setMouseTracking(True)
        self.scaled_pixmap = None
        self.original_pixmap = None
        self.scale_factor = 1.0
        self.zoom_level = 1.0  # Additional zoom level for user zooming
        self.offset = QPoint(0, 0)
        self.page_configs = []  # Direct access to page_configs

        # Drawing variables
        self.drawing = False
        self.start_pos = None
        self.current_pos = None
        self.drawing_rect = None
        self.drawing_line = None
        self.drawing_mode = None  # 'region' or 'column'
        self.current_region_type = None  # 'header', 'items', or 'summary'
        self.current_region_index = None  # Index of the region being edited

        # Hover variables for delete icons
        self.hover_region_type = None
        self.hover_region_index = None
        self.hover_column_type = None
        self.hover_column_index = None
        self.hover_delete_icon = False
        self.delete_icon_size = 16  # Size of the delete icon in pixels

        # Set cursor
        self.normal_cursor = Qt.ArrowCursor
        self.drawing_cursor = Qt.CrossCursor
        self.column_cursor = Qt.SplitHCursor
        self.setCursor(self.normal_cursor)

        # Enable wheel events for zooming
        self.setFocusPolicy(Qt.StrongFocus)

    def setPixmap(self, pixmap):
        self.original_pixmap = pixmap  # Store the original pixmap
        super().setPixmap(pixmap)
        self.adjustPixmap()

    def wheelEvent(self, event):
        """Handle mouse wheel events for zooming"""
        try:
            # Check if we have a pixmap to zoom
            if not self.original_pixmap:
                event.ignore()
                return

            # Get the delta and determine zoom direction
            delta = event.angleDelta().y()

            # Use smaller zoom factors for smoother zooming
            zoom_factor = 1.05 if delta > 0 else 0.95

            # Get current zoom level before changing
            old_zoom = self.zoom_level

            # Apply zoom
            self.zoom_level *= zoom_factor

            # Stricter zoom range (0.5x to 2.5x) to prevent scaling issues
            self.zoom_level = max(0.5, min(2.5, self.zoom_level))

            # Only proceed if zoom actually changed
            if abs(old_zoom - self.zoom_level) > 0.01:
                # Adjust the pixmap with the new zoom level
                self.adjustPixmap()

                # Emit signal with new zoom level
                try:
                    self.zoom_changed.emit(self.zoom_level)
                    print(f"[DEBUG] Emitted zoom_changed signal with level: {self.zoom_level:.2f}x")
                except Exception as e:
                    print(f"[DEBUG] Error emitting zoom_changed signal: {str(e)}")
                    # If signal emission fails, try to update the parent's zoom label directly
                    if self.parent and hasattr(self.parent, 'update_zoom_label'):
                        self.parent.update_zoom_label(self.zoom_level)
                        print(f"[DEBUG] Called parent's update_zoom_label directly with level: {self.zoom_level:.2f}x")

                # Ensure proper scrolling after zoom
                if self.parent and hasattr(self.parent, 'ensure_full_scroll_range'):
                    self.parent.ensure_full_scroll_range()

            # Accept the event
            event.accept()
        except Exception as e:
            print(f"[DEBUG] Error in PDFLabel.wheelEvent: {str(e)}")
            import traceback
            traceback.print_exc()
            event.ignore()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Prevent recursion by checking if we're already in adjustPixmap
        if not hasattr(self, '_in_adjust_pixmap') or not self._in_adjust_pixmap:
            self.adjustPixmap()

    def adjustPixmap(self):
        """Adjust the pixmap based on the current zoom level and scale factor"""
        try:
            # Check if we have a pixmap to adjust
            if not self.pixmap():
                print(f"[DEBUG] No pixmap to adjust")
                return

            # Check if parent exists and if the PDF section is hidden
            if self.parent and hasattr(self.parent, 'pdf_section_was_hidden') and self.parent.pdf_section_was_hidden:
                print(f"[DEBUG] Skipping adjustPixmap because PDF section is hidden")
                return

            # Set recursion prevention flag
            self._in_adjust_pixmap = True

            # Store the original pixmap if not already stored
            if not hasattr(self, 'original_pixmap') or self.original_pixmap is None:
                self.original_pixmap = self.pixmap()
                print(f"[DEBUG] Stored original pixmap of size {self.original_pixmap.size().width()}x{self.original_pixmap.size().height()}")

            # Calculate scaling to fit the label while maintaining aspect ratio
            label_size = self.size()
            pixmap_size = self.original_pixmap.size()

            # A4 size in pixels at 72 DPI: 595 x 842 (portrait)
            a4_width = 595
            a4_height = 842

            # Calculate scaling to display at A4 size
            a4_scale_x = a4_width / pixmap_size.width()
            a4_scale_y = a4_height / pixmap_size.height()
            a4_scale = min(a4_scale_x, a4_scale_y)

            # Calculate scaling to fit the label
            fit_scale_x = label_size.width() / pixmap_size.width()
            fit_scale_y = label_size.height() / pixmap_size.height()
            fit_scale = min(fit_scale_x, fit_scale_y)

            # Use A4 scale unless it's too big for the label, then use fit scale
            if a4_scale > fit_scale:
                self.scale_factor = max(0.1, fit_scale)  # Ensure minimum scale factor
                print(f"[DEBUG] Using fit scale: {self.scale_factor:.2f} (A4 scale {a4_scale:.2f} too large)")
            else:
                self.scale_factor = a4_scale
                print(f"[DEBUG] Using A4 scale: {self.scale_factor:.2f}")

            # Apply the user's zoom level
            if not hasattr(self, 'zoom_level'):
                self.zoom_level = 1.0

            # Ensure zoom level stays within reasonable limits
            self.zoom_level = max(0.5, min(3.0, self.zoom_level))

            # Calculate effective scale (base scale * zoom level)
            effective_scale = self.scale_factor * self.zoom_level

            # Ensure minimum effective scale to prevent excessive debug messages
            if effective_scale < 0.1:
                effective_scale = 0.1

            # Calculate the scaled size
            scaled_width = max(100, int(pixmap_size.width() * effective_scale))
            scaled_height = max(100, int(pixmap_size.height() * effective_scale))

            # Create a scaled version of the pixmap
            self.scaled_pixmap = self.original_pixmap.scaled(
                scaled_width, scaled_height,
                Qt.KeepAspectRatio, Qt.SmoothTransformation
            )

            # Update the displayed pixmap
            super().setPixmap(self.scaled_pixmap)
            print(f"[DEBUG] Set scaled pixmap of size {scaled_width}x{scaled_height} (zoom: {self.zoom_level:.2f}x)")

            # Set the size of the label to match the scaled pixmap size
            # This allows scrolling when the image is larger than the viewport
            # Add extra padding at the bottom to ensure the entire document is visible
            # Add more padding when zoomed in (proportional to zoom level)
            extra_padding = int(300 * self.zoom_level)  # More padding at higher zoom levels
            self.setFixedSize(scaled_width, scaled_height + extra_padding)

            # Calculate offset to center the image if it's smaller than the viewport
            if scaled_width < label_size.width():
                self.offset = QPoint((label_size.width() - scaled_width) // 2, 0)
            else:
                self.offset = QPoint(0, 0)

            if scaled_height < label_size.height():
                self.offset.setY((label_size.height() - scaled_height) // 2)
            else:
                self.offset.setY(0)

            # Update the display
            self.update()

            # If parent has ensure_full_scroll_range method, call it
            if self.parent and hasattr(self.parent, 'ensure_full_scroll_range'):
                self.parent.ensure_full_scroll_range()

            # Clear recursion prevention flag
            self._in_adjust_pixmap = False
        except Exception as e:
            print(f"[DEBUG] Error in adjustPixmap: {str(e)}")
            import traceback
            traceback.print_exc()
            # Clear recursion prevention flag even if there's an error
            self._in_adjust_pixmap = False

    def mapToPixmap(self, pos):
        """Convert widget coordinates to pixmap coordinates"""
        if not self.scaled_pixmap:
            return pos

        try:
            # Check if pos is a valid QPoint
            if not hasattr(pos, 'x') or not hasattr(pos, 'y'):
                print(f"[DEBUG] Invalid position object in mapToPixmap: {pos}")
                return QPoint(0, 0)

            # With scrolling support, we don't need to adjust for offset
            # Just check if the point is within the scaled pixmap
            if (pos.x() < 0 or pos.y() < 0 or
                pos.x() >= self.scaled_pixmap.width() or
                pos.y() >= self.scaled_pixmap.height()):
                return QPoint(-1, -1)  # Out of bounds

            # Calculate effective scale factor (base scale * zoom level)
            effective_scale = self.scale_factor * self.zoom_level

            # Convert to original pixmap coordinates
            x = pos.x() / effective_scale
            y = pos.y() / effective_scale

            return QPoint(int(x), int(y))
        except Exception as e:
            print(f"[DEBUG] Error in mapToPixmap: {str(e)}")
            return QPoint(0, 0)

    def mapFromPixmap(self, pos):
        """Convert pixmap coordinates to widget coordinates"""
        if not self.scaled_pixmap:
            return pos

        try:
            # Check if pos is a valid QPoint
            if not hasattr(pos, 'x') or not hasattr(pos, 'y'):
                print(f"[DEBUG] Invalid position object in mapFromPixmap: {pos}")
                return QPoint(0, 0)

            # Calculate effective scale factor (base scale * zoom level)
            effective_scale = self.scale_factor * self.zoom_level

            # Scale the coordinates
            x = pos.x() * effective_scale
            y = pos.y() * effective_scale

            # With scrolling support, we don't need to add offset
            return QPoint(int(x), int(y))
        except Exception as e:
            print(f"[DEBUG] Error in mapFromPixmap: {str(e)}")
            return QPoint(0, 0)

    def mousePressEvent(self, event):
        if self.parent and event.button() == Qt.LeftButton:
            pos = event.pos()

            # Check if we're clicking on a delete icon
            if self.hover_delete_icon:
                if self.hover_region_type is not None and self.hover_region_index is not None:
                    # Delete the region
                    self.parent.delete_region(self.hover_region_type, self.hover_region_index)
                    # Reset hover states
                    self.hover_region_type = None
                    self.hover_region_index = None
                    self.hover_delete_icon = False
                    self.setCursor(self.normal_cursor)
                    return
                elif self.hover_column_type is not None and self.hover_column_index is not None:
                    # Delete the column
                    self.parent.delete_column(self.hover_column_type, self.hover_column_index)
                    # Reset hover states
                    self.hover_column_type = None
                    self.hover_column_index = None
                    self.hover_delete_icon = False
                    self.setCursor(self.normal_cursor)
                    return

            # If not clicking on a delete icon, proceed with normal handling
            # No need to check offset with scrolling support
            self.parent.handle_mouse_press(self.mapToPixmap(pos))

    def mouseMoveEvent(self, event):
        if self.parent:
            pos = event.pos()
            pixmap_pos = self.mapToPixmap(pos)
            self.parent.handle_mouse_move(pixmap_pos)

            # Check if we're not already drawing
            if not self.drawing:
                # Reset hover states
                old_hover_region_type = self.hover_region_type
                old_hover_region_index = self.hover_region_index
                old_hover_column_type = self.hover_column_type
                old_hover_column_index = self.hover_column_index
                old_hover_delete_icon = self.hover_delete_icon

                self.hover_region_type = None
                self.hover_region_index = None
                self.hover_column_type = None
                self.hover_column_index = None
                self.hover_delete_icon = False

                # Check if we're hovering over a region's delete icon
                if hasattr(self.parent, 'regions'):
                    for region_type, region_list in self.parent.regions.items():
                        for i, region in enumerate(region_list):
                            # Use standardized coordinate system - single format everywhere
                            from standardized_coordinates import StandardRegion

                            # Enforce StandardRegion format - NO backward compatibility
                            if not isinstance(region, StandardRegion):
                                print(f"[ERROR] Invalid region type in {region_type}[{i}]: expected StandardRegion, got {type(region)}")
                                continue

                            rect = region.rect

                            # Convert pixmap coordinates to widget coordinates
                            scaled_rect = QRect(
                                self.mapFromPixmap(QPoint(rect.x(), rect.y())),
                                self.mapFromPixmap(QPoint(rect.x() + rect.width(), rect.y() + rect.height()))
                            )

                            # Create delete icon rect in the top-right corner
                            delete_icon_rect = QRect(
                                scaled_rect.right() - self.delete_icon_size - 2,
                                scaled_rect.top() + 2,
                                self.delete_icon_size,
                                self.delete_icon_size
                            )

                            # Check if mouse is over the delete icon
                            if delete_icon_rect.contains(pos):
                                self.hover_region_type = region_type
                                self.hover_region_index = i
                                self.hover_delete_icon = True
                                self.setCursor(Qt.PointingHandCursor)
                                break

                            # Check if mouse is over the region (for showing delete icon)
                            elif scaled_rect.contains(pos):
                                self.hover_region_type = region_type
                                self.hover_region_index = i
                                self.setCursor(self.normal_cursor)
                                break

                        # Break out of outer loop if we found a match
                        if self.hover_region_type is not None:
                            break

                # If not hovering over a region, check if we're hovering over a column's delete icon
                if self.hover_region_type is None and hasattr(self.parent, 'column_lines'):
                    for region_type, lines in self.parent.column_lines.items():
                        # Store the original region_type (enum or string) for later use
                        original_region_type = region_type

                        for column_number, line in enumerate(lines):
                            # Get start and end points of the column line
                            if isinstance(line, tuple) and len(line) >= 2:
                                start_point = line[0]
                                end_point = line[1]
                            elif isinstance(line, list) and len(line) >= 2:
                                start_point = line[0]
                                end_point = line[1]

                            # Convert to widget coordinates
                            start = self.mapFromPixmap(start_point)
                            end = self.mapFromPixmap(end_point)

                            # Position the delete icon just above the column line
                            icon_x = start.x()
                            icon_y = start.y() - 25  # Position above the column line

                            # Create delete icon rect above the column line
                            delete_icon_rect = QRect(
                                icon_x - self.delete_icon_size // 2,
                                icon_y - self.delete_icon_size // 2,
                                self.delete_icon_size,
                                self.delete_icon_size
                            )

                            # Check if mouse is over the delete icon
                            if delete_icon_rect.contains(pos):
                                self.hover_column_type = original_region_type
                                self.hover_column_index = column_number
                                self.hover_delete_icon = True
                                self.setCursor(Qt.PointingHandCursor)
                                print(f"[DEBUG] Hovering over delete icon for column {column_number} of type {original_region_type}")
                                break

                            # Check if mouse is near the top of the column line (for showing delete icon)
                            # Use a small tolerance area around the top of the line
                            elif abs(pos.x() - start.x()) < 10 and start.y() - 30 <= pos.y() <= start.y() + 30:
                                self.hover_column_type = original_region_type
                                self.hover_column_index = column_number
                                self.setCursor(self.normal_cursor)
                                print(f"[DEBUG] Hovering near column {column_number} of type {original_region_type}")
                                break

                        # Break out of outer loop if we found a match
                        if self.hover_column_type is not None:
                            break

                # If hover state changed, update the display
                if (old_hover_region_type != self.hover_region_type or
                    old_hover_region_index != self.hover_region_index or
                    old_hover_column_type != self.hover_column_type or
                    old_hover_column_index != self.hover_column_index or
                    old_hover_delete_icon != self.hover_delete_icon):
                    self.update()

                # If not hovering over a delete icon, use auto-detect regions for column drawing
                if not self.hover_delete_icon and hasattr(self.parent, 'auto_detect_regions'):
                    # Auto-detect if we're hovering over a region for column drawing
                    self.parent.auto_detect_regions(pixmap_pos)

    def keyPressEvent(self, event):
        """Handle key press events

        Allows deleting regions and columns using the Delete key when hovering over them.
        This provides an alternative to clicking on the delete icon (X).
        """
        # Check if Delete key is pressed
        if event.key() == Qt.Key_Delete:
            # Check if there's an active region being hovered over
            if self.hover_region_type is not None and self.hover_region_index is not None:
                # Delete the region
                self.parent.delete_region(self.hover_region_type, self.hover_region_index)
                # Reset hover states
                self.hover_region_type = None
                self.hover_region_index = None
                self.hover_delete_icon = False
                self.setCursor(self.normal_cursor)
                print(f"[DEBUG] Deleted region using Delete key")
                event.accept()
                return
            # Check if there's an active column being hovered over
            elif self.hover_column_type is not None and self.hover_column_index is not None:
                # Delete the column
                self.parent.delete_column(self.hover_column_type, self.hover_column_index)
                # Reset hover states
                self.hover_column_type = None
                self.hover_column_index = None
                self.hover_delete_icon = False
                self.setCursor(self.normal_cursor)
                print(f"[DEBUG] Deleted column using Delete key")
                event.accept()
                return

        # Pass unhandled events to parent
        super().keyPressEvent(event)

    def mouseReleaseEvent(self, event):
        if self.parent and event.button() == Qt.LeftButton:
            pos = event.pos()
            self.parent.handle_mouse_release(self.mapToPixmap(pos))

    def paintEvent(self, event):
        if not self.scaled_pixmap:
            super().paintEvent(event)
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        # Draw the scaled pixmap at the current position (with scrolling support)
        painter.drawPixmap(0, 0, self.scaled_pixmap)

        # Draw regions if parent has them
        if self.parent and hasattr(self.parent, 'regions'):
            # Get colors from parent's theme if available
            if hasattr(self.parent, 'theme'):
                colors = {
                    'header': QColor(52, 152, 219, 127),  # Blue with transparency
                    'items': QColor(46, 204, 113, 127),   # Green with transparency
                    'summary': QColor(155, 89, 182, 127)  # Purple with transparency
                }
            else:
                # Fallback colors
                colors = {
                    'header': QColor(52, 152, 219, 127),  # Blue with transparency
                    'items': QColor(46, 204, 113, 127),   # Green with transparency
                    'summary': QColor(155, 89, 182, 127)  # Purple with transparency
                }

            # Draw regions from parent.regions
            for region_type, region_list in self.parent.regions.items():
                color = colors.get(region_type, QColor(200, 200, 200))  # Default gray
                pen = QPen(color, 2, Qt.SolidLine)
                painter.setPen(pen)

                for i, region in enumerate(region_list):
                    # Use standardized coordinate system - single format everywhere
                    from standardized_coordinates import StandardRegion

                    # Enforce StandardRegion format - NO backward compatibility
                    if not isinstance(region, StandardRegion):
                        log_error(f"Invalid region type in {region_type}[{i}]: expected StandardRegion, got {type(region)}")
                        continue

                    rect = region.rect
                    custom_label = region.label

                    # Convert pixmap coordinates to widget coordinates
                    scaled_rect = QRect(
                        self.mapFromPixmap(QPoint(rect.x(), rect.y())),
                        self.mapFromPixmap(QPoint(rect.x() + rect.width(), rect.y() + rect.height()))
                    )

                    # If drawing a column and this is the active rectangle, use a stronger fill
                    if (hasattr(self.parent, 'drawing_column') and self.parent.drawing_column and
                        hasattr(self.parent, 'active_rect_index') and hasattr(self.parent, 'active_region_type') and
                        region_type == self.parent.active_region_type and i == self.parent.active_rect_index):
                        # Highlight the active rectangle with a semi-transparent fill
                        painter.fillRect(scaled_rect, QColor(color.red(), color.green(), color.blue(), 50))
                        print(f"[DEBUG] Highlighting active rectangle for column drawing: {region_type} {i}")

                    # Draw the rectangle
                    painter.drawRect(scaled_rect)

                    # Draw label text on the left side of the rectangle
                    painter.setPen(Qt.black)
                    font = QFont("Arial", 10, QFont.Bold)
                    painter.setFont(font)

                    # Use the custom label if available, otherwise create one
                    if custom_label:
                        label_text = custom_label
                    else:
                        # Create short label with section type and table number
                        titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
                        label_text = titles.get(region_type, region_type[0].upper())
                        label_text += str(i+1)  # Add table number

                    # Position the label on the left side of the rectangle
                    label_x = scaled_rect.left() - 30  # 30px to the left of the rectangle
                    label_y = scaled_rect.top() + scaled_rect.height() // 2  # Center vertically

                    # Create label background
                    label_bg_rect = QRect(label_x, label_y - 10, 25, 20)
                    label_bg_color = QColor(255, 255, 255, 200)
                    painter.fillRect(label_bg_rect, label_bg_color)

                    # Draw border around label background
                    painter.setPen(Qt.black)
                    painter.drawRect(label_bg_rect)

                    # Draw connecting line from label to rectangle
                    connecting_line_start = QPoint(label_x + 25, label_y)
                    connecting_line_end = QPoint(scaled_rect.left(), label_y)
                    painter.drawLine(connecting_line_start, connecting_line_end)

                    # Draw the label text
                    painter.drawText(label_bg_rect, Qt.AlignCenter, label_text)

                    # Draw delete icon (X) in the top-right corner when hovering
                    if self.hover_region_type == region_type and self.hover_region_index == i:
                        delete_icon_size = self.delete_icon_size
                        delete_icon_rect = QRect(
                            scaled_rect.right() - delete_icon_size - 2,
                            scaled_rect.top() + 2,
                            delete_icon_size,
                            delete_icon_size
                        )

                        # Draw white background circle for the X
                        painter.setBrush(QColor(255, 255, 255, 220))
                        painter.setPen(QPen(Qt.red, 1))
                        painter.drawEllipse(delete_icon_rect)

                        # Draw X
                        painter.setPen(QPen(Qt.red, 2))
                        painter.drawLine(
                            delete_icon_rect.left() + 3,
                            delete_icon_rect.top() + 3,
                            delete_icon_rect.right() - 3,
                            delete_icon_rect.bottom() - 3
                        )
                        painter.drawLine(
                            delete_icon_rect.right() - 3,
                            delete_icon_rect.top() + 3,
                            delete_icon_rect.left() + 3,
                            delete_icon_rect.bottom() - 3
                        )

                    # Reset pen color for next rectangle
                    painter.setPen(pen)

        # Draw column lines if parent has them
        if self.parent and hasattr(self.parent, 'column_lines'):
            for region_type, lines in self.parent.column_lines.items():
                # Get color based on region type
                color = colors.get(region_type.value if hasattr(region_type, 'value') else region_type, QColor(200, 200, 200))
                pen = QPen(color, 2, Qt.DashLine)
                painter.setPen(pen)

                # Draw each column line
                for column_number, line in enumerate(lines):
                    # Get start and end points
                    if isinstance(line, tuple) and len(line) >= 2:
                        start_point = line[0]
                        end_point = line[1]

                        # Get region index if available
                        region_index = line[2] if len(line) > 2 else None

                        # Print debug info
                        if region_index is not None:
                            print(f"[DEBUG] Drawing column line {column_number+1} for region index {region_index}")
                    elif isinstance(line, list) and len(line) >= 2:
                        # Handle list format
                        start_point = line[0]
                        end_point = line[1]
                        region_index = line[2] if len(line) > 2 else None
                    else:
                        # Unexpected format, skip this line
                        print(f"[DEBUG] Skipping column line with unexpected format: {line}")
                        continue

                    # Convert to widget coordinates
                    start = self.mapFromPixmap(start_point)
                    end = self.mapFromPixmap(end_point)

                    # Draw the line
                    painter.drawLine(start, end)

                    # Draw the column number label
                    painter.setPen(Qt.black)
                    font = QFont("Arial", 8, QFont.Bold)
                    painter.setFont(font)
                    label_text = f"C{column_number+1}"

                    # Create label with background
                    label_x = start.x() + 2
                    label_y = start.y() - 15
                    label_bg_rect = QRect(label_x, label_y, 20, 15)
                    label_bg_color = QColor(255, 255, 255, 200)
                    painter.fillRect(label_bg_rect, label_bg_color)

                    # Draw the label text
                    painter.drawText(label_bg_rect, Qt.AlignCenter, label_text)

                    # Draw delete icon (X) in the middle of the column when hovering
                    # Compare using string representation for enum types
                    hover_type_str = str(self.hover_column_type) if self.hover_column_type is not None else None
                    region_type_str = str(region_type) if region_type is not None else None

                    if ((self.hover_column_type == region_type or hover_type_str == region_type_str) and
                        self.hover_column_index == column_number):
                        delete_icon_size = self.delete_icon_size

                        # Position the delete icon just above the column line
                        icon_x = start.x()
                        icon_y = start.y() - 25  # Position above the column line

                        delete_icon_rect = QRect(
                            icon_x - delete_icon_size // 2,
                            icon_y - delete_icon_size // 2,
                            delete_icon_size,
                            delete_icon_size
                        )

                        # Draw white background circle for the X
                        painter.setBrush(QColor(255, 255, 255, 220))
                        painter.setPen(QPen(Qt.red, 1))
                        painter.drawEllipse(delete_icon_rect)

                        # Draw X
                        painter.setPen(QPen(Qt.red, 2))
                        painter.drawLine(
                            delete_icon_rect.left() + 3,
                            delete_icon_rect.top() + 3,
                            delete_icon_rect.right() - 3,
                            delete_icon_rect.bottom() - 3
                        )
                        painter.drawLine(
                            delete_icon_rect.right() - 3,
                            delete_icon_rect.top() + 3,
                            delete_icon_rect.left() + 3,
                            delete_icon_rect.bottom() - 3
                        )

                    # Reset pen for next line
                    painter.setPen(pen)

        # Draw the current rectangle being drawn
        if self.drawing and self.start_pos and self.current_pos:
            if self.drawing_mode == 'region':
                # Draw the rectangle being drawn
                start_widget = self.mapFromPixmap(self.start_pos)
                current_widget = self.mapFromPixmap(self.current_pos)

                # Get color based on current region type
                color = colors.get(self.current_region_type, QColor(200, 200, 200))
                pen = QPen(color, 2, Qt.SolidLine)
                painter.setPen(pen)

                # Calculate rectangle
                rect = QRect(start_widget, current_widget).normalized()
                painter.drawRect(rect)

                # Draw label for the rectangle being drawn
                painter.setPen(Qt.black)
                font = QFont("Arial", 10, QFont.Bold)
                painter.setFont(font)

                # Create short label with section type
                titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
                label_text = titles.get(self.current_region_type, self.current_region_type[0].upper())

                # Position and draw the label on the left side of the rectangle being drawn
                label_x = rect.left() - 30
                label_y = rect.top() + rect.height() // 2

                # Create label background
                label_bg_rect = QRect(label_x, label_y - 10, 25, 20)
                label_bg_color = QColor(255, 255, 255, 200)
                painter.fillRect(label_bg_rect, label_bg_color)

                # Draw border around label background
                painter.setPen(Qt.black)
                painter.drawRect(label_bg_rect)

                # Draw connecting line from label to rectangle
                connecting_line_start = QPoint(label_x + 25, label_y)
                connecting_line_end = QPoint(rect.left(), label_y)
                painter.drawLine(connecting_line_start, connecting_line_end)

                # Draw the label text
                painter.drawText(label_bg_rect, Qt.AlignCenter, label_text)

            elif self.drawing_mode == 'column':
                # Draw the column line being drawn
                # Check if parent has current_rect (the active rectangle)
                if hasattr(self.parent, 'current_rect') and self.parent.current_rect:
                    rect = self.parent.current_rect

                    # Use a dashed blue line for column preview
                    pen = QPen(QColor(65, 105, 225), 2, Qt.DashLine)
                    painter.setPen(pen)

                    # Keep x position within the rectangle boundaries
                    x_pos = max(rect.left(), min(self.current_pos.x(), rect.right()))

                    # Draw vertical line from top to bottom of the rectangle
                    start = self.mapFromPixmap(QPoint(x_pos, rect.top()))
                    end = self.mapFromPixmap(QPoint(x_pos, rect.bottom()))

                    print(f"[DEBUG] Drawing preview column line at x={x_pos} in rectangle {rect.x()},{rect.y()},{rect.width()},{rect.height()}")
                    painter.drawLine(start, end)
                else:
                    # No active rectangle, don't draw anything
                    print(f"[DEBUG] No active rectangle for column drawing")

class SplitScreenInvoiceProcessor(QWidget):
    """
    A split-screen widget that combines PDF drawing functionality on the left
    and extraction results on the right, with a draggable divider.
    """

    # Define signals
    region_drawn = Signal()  # Signal to indicate a region was drawn
    column_line_drawn = Signal()  # Signal to indicate a column line was drawn
    go_back = Signal()  # Signal to go back to previous screen
    save_template_signal = Signal()  # Signal to save template
    config_completed = Signal(dict)  # Signal to indicate config is completed
    invoice2data_template_created = Signal(dict)  # Signal to indicate invoice2data template was created

    def __init__(self, parent=None):
        super().__init__(parent)
        self.pdf_path = None
        self.pdf_document = None
        self.current_page_index = 0

        # Import standardized coordinate system - Convert Once, Use Everywhere
        from standardized_coordinates import StandardRegion

        # Initialize database connection
        from database import InvoiceDatabase
        self.db = InvoiceDatabase()

        # Use standardized coordinate format throughout
        self.regions = {'header': [], 'items': [], 'summary': []}  # Will contain StandardRegion objects
        self.column_lines = {'header': [], 'items': [], 'summary': []}  # Standardized keys
        self.table_areas = {}
        # Initialize extraction parameters with section-specific structure only
        self.extraction_params = {
            'header': {'row_tol': 5, 'flavor': 'stream', 'split_text': True, 'strip_text': '\n'},
            'items': {'row_tol': 15, 'flavor': 'stream', 'split_text': True, 'strip_text': '\n'},
            'summary': {'row_tol': 10, 'flavor': 'stream', 'split_text': True, 'strip_text': '\n'}
        }

        # Invoice2data template variables
        self.invoice2data_template = {
            "issuer": "",
            "fields": {},
            "lines": {
                "start": "",
                "end": "",
                "first_line": [],
                "line": "",
                "types": {}
            },
            "keywords": [],
            "options": {
                "currency": "INR",
                "languages": ["en"],
                "decimal_separator": ".",
                "date_formats": [],  # Empty by default, user can add if needed
                "remove_whitespace": False,
                "remove_accents": False,
                "lowercase": False,
                "replace": []  # Empty by default, user can add if needed
            }
        }

        # Drawing state variables
        self.current_region_type = None  # Current region type being drawn
        self.drawing_column = False  # Whether we're drawing a column line
        self.active_region_type = None  # Active region type for editing
        self.active_rect_index = None  # Active rectangle index for editing

        # Auto-switching to column drawing mode
        self.auto_column_mode = False  # Whether to auto-switch to column drawing mode (off by default)
        self.hover_region_type = None  # Region type being hovered over
        self.hover_rect_index = None  # Rectangle index being hovered over
        self.last_cursor_update = 0  # Time of last cursor update (to avoid too frequent updates)

        # Undo feature has been removed as per user preference
        # But we still need to initialize the undo_stack attribute to prevent errors
        self.undo_stack = []

        # Define theme colors
        self.theme = {
            "primary": "#4169E1",       # Royal Blue
            "primary_dark": "#3159C1",  # Darker blue
            "secondary": "#28a745",     # Green
            "tertiary": "#8B5CF6",      # Violet
            "danger": "#D32F2F",        # Red
            "warning": "#F59E0B",       # Amber
            "light": "#F9FAFB",         # Light gray
            "dark": "#111827",          # Dark gray
            "bg": "#F3F4F6",            # Background light gray
            "text": "#1F2937",          # Text dark
            "border": "#E5E7EB",        # Border light gray
        }

        # Multi-page support
        self.multi_page_mode = False
        self.page_regions = {}
        self.page_column_lines = {}
        self.page_configs = []
        self.use_middle_page = False
        self.fixed_page_count = 0

        # Coordinate conversion parameters for standardized system
        self.scale_x = 1.0
        self.scale_y = 1.0
        self.page_height = 842.0  # Default A4 height

        # Initialize cached extraction data
        self._cached_extraction_data = {
            'header': [],
            'items': [],
            'summary': []
        }
        self._last_extraction_state = None

        # Flag to skip automatic extraction update after specific region extraction
        self._skip_extraction_update = False

        # Track whether region labels have been set for each page
        self._region_labels_set = {}  # Dictionary of page_index -> {section_type -> {region_index -> bool}}

        # Track whether metadata has been set
        self._metadata_set = False

        # Initialize extraction method
        self.current_extraction_method = "pypdf_table_extraction"

        # Initialize UI
        self.initUI()

        # Set up keyboard shortcuts
        self.setup_shortcuts()

        # Update scale factors when PDF is loaded
        self.update_coordinate_scale_factors()

    def on_extraction_method_changed(self, method):
        """Handle extraction method change"""
        self.current_extraction_method = method
        print(f"[DEBUG] Extraction method changed to: {method}")

        # Update extraction parameters based on method
        if method == "pdftotext":
            print("[INFO] Using pdftotext extraction method")
        elif method == "tesseract_ocr":
            print("[INFO] Using tesseract OCR extraction method")
        else:
            print("[INFO] Using pypdf_table_extraction method")

        # Trigger re-extraction if we have regions defined
        if self.regions and any(self.regions.values()):
            self.update_extraction_results(force=True)

    def update_coordinate_scale_factors(self):
        """Update coordinate scale factors for standardized coordinate conversion"""
        try:
            if self.pdf_path and self.pdf_document:
                from pdf_extraction_utils import get_scale_factors
                scale_info = get_scale_factors(self.pdf_path, self.current_page_index)
                self.scale_x = scale_info.get('scale_x', 1.0)
                self.scale_y = scale_info.get('scale_y', 1.0)

                # Get page height for coordinate conversion
                page = self.pdf_document[self.current_page_index]
                self.page_height = page.mediabox.height

                print(f"[DEBUG] Updated coordinate scale factors: scale_x={self.scale_x:.4f}, scale_y={self.scale_y:.4f}, page_height={self.page_height}")
        except Exception as e:
            print(f"[WARNING] Failed to update coordinate scale factors: {e}")
            # Use defaults
            self.scale_x = 1.0
            self.scale_y = 1.0
            self.page_height = 842.0

    def create_standard_region(self, x: int, y: int, width: int, height: int, region_type: str, region_index: int = None) -> 'StandardRegion':
        """Create a standardized region from UI coordinates - SINGLE ENTRY POINT

        This is the ONLY method that should be used to create regions from UI input.
        It ensures all regions are in standardized format from creation.

        Args:
            x, y, width, height: UI coordinates from mouse drawing
            region_type: 'header', 'items', or 'summary'
            region_index: Optional index for label generation

        Returns:
            StandardRegion with both UI and extraction coordinates
        """
        try:
            from coordinate_boundary_converters import UIInputConverter

            # Generate label
            titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
            if region_index is None:
                region_index = len(self.regions.get(region_type, []))
            label = f"{titles.get(region_type, 'R')}{region_index + 1}"

            # Create standardized region with both coordinate systems
            standard_region = UIInputConverter.from_mouse_drawing(
                x, y, width, height, label,
                self.scale_x, self.scale_y, self.page_height
            )

            print(f"[DEBUG] Created standard region {label}: UI({x},{y},{width},{height}) -> Extraction{standard_region.extraction_coords}")
            return standard_region

        except Exception as e:
            print(f"[ERROR] Failed to create standard region: {e}")
            raise

    def get_region_color(self, region_type):
        """Get the color for a region type"""
        colors = {
            'header': QColor(52, 152, 219, 127),  # Blue with transparency
            'items': QColor(46, 204, 113, 127),   # Green with transparency
            'summary': QColor(155, 89, 182, 127)  # Purple with transparency
        }

        # Handle both string and enum types
        if hasattr(region_type, 'value'):
            return colors.get(region_type.value, QColor(200, 200, 200, 127))
        else:
            return colors.get(region_type, QColor(200, 200, 200, 127))

    def setup_shortcuts(self):
        """Set up keyboard shortcuts

        This sets up the Delete key shortcut for deleting regions and columns.
        Note: The Delete key also works directly in the PDFLabel widget when hovering over regions or columns.
        """
        # Delete key for deleting the active drawing
        delete_shortcut = QShortcut(QKeySequence(Qt.Key_Delete), self)
        delete_shortcut.activated.connect(self.delete_active_drawing)

    def delete_active_drawing(self):
        """Delete the active drawing (region or column) that's currently being hovered over

        This method is called when the Delete key shortcut is activated.
        It provides the same functionality as pressing Delete when hovering over a region or column.
        """
        if not self.pdf_document:
            return

        # Check if there's an active region being hovered over
        if (hasattr(self.pdf_label, 'hover_region_type') and self.pdf_label.hover_region_type is not None and
            hasattr(self.pdf_label, 'hover_region_index') and self.pdf_label.hover_region_index is not None):
            # Delete the region
            self.delete_region(self.pdf_label.hover_region_type, self.pdf_label.hover_region_index)
            # Reset hover states
            self.pdf_label.hover_region_type = None
            self.pdf_label.hover_region_index = None
            self.pdf_label.hover_delete_icon = False
            self.pdf_label.setCursor(self.pdf_label.normal_cursor)
            print(f"[DEBUG] Deleted region using Delete key shortcut")
        # Check if there's an active column being hovered over
        elif (hasattr(self.pdf_label, 'hover_column_type') and self.pdf_label.hover_column_type is not None and
              hasattr(self.pdf_label, 'hover_column_index') and self.pdf_label.hover_column_index is not None):
            # Delete the column
            self.delete_column(self.pdf_label.hover_column_type, self.pdf_label.hover_column_index)
            # Reset hover states
            self.pdf_label.hover_column_type = None
            self.pdf_label.hover_column_index = None
            self.pdf_label.hover_delete_icon = False
            self.pdf_label.setCursor(self.pdf_label.normal_cursor)
            print(f"[DEBUG] Deleted column using Delete key shortcut")

    def _update_table_areas_after_region_deletion(self, region_type, region_index):
        """Update table_areas dictionary after a region is deleted

        Args:
            region_type (str or RegionType): The type of region that was deleted
            region_index (int): The index of the region that was deleted
        """
        # Convert region_type to string if it's an enum
        region_type_str = region_type.value if hasattr(region_type, 'value') else region_type

        # Check if we have table_areas attribute
        if hasattr(self, '_table_areas'):
            # If the region type exists in table_areas, remove the corresponding entry
            if region_type_str in self._table_areas and isinstance(self._table_areas[region_type_str], list):
                # Check if the index is valid
                if 0 <= region_index < len(self._table_areas[region_type_str]):
                    # Remove the table area
                    self._table_areas[region_type_str].pop(region_index)
                    print(f"[DEBUG] Removed table area for {region_type_str} at index {region_index}")

                    # Update indices for regions with higher indices
                    for i in range(region_index, len(self._table_areas[region_type_str])):
                        print(f"[DEBUG] Updating table area index from {i+1} to {i} for {region_type_str}")

                    # If there are no more table areas for this region type, initialize an empty list
                    if not self._table_areas[region_type_str]:
                        self._table_areas[region_type_str] = []
        else:
            # Initialize _table_areas if it doesn't exist
            self._table_areas = {
                'header': [],
                'items': [],
                'summary': []
            }

        print(f"[DEBUG] Updated table_areas after region deletion")

    def _update_extraction_results_after_region_deletion(self, region_type_str, deleted_region_index):
        """Update extraction results after a region is deleted to ensure region labels are consistent

        Args:
            region_type_str (str): The type of region that was deleted (header, items, summary)
            deleted_region_index (int): The index of the region that was deleted
        """
        print(f"[DEBUG] Updating extraction results after deletion of {region_type_str} region at index {deleted_region_index}")

        # Get the current data
        current_data = self._get_current_json_data()
        if not current_data or region_type_str not in current_data or current_data[region_type_str] is None:
            print(f"[DEBUG] No extraction data to update for {region_type_str}")
            return

        # Get the data for this section
        section_data = current_data[region_type_str]

        # Handle different data types
        if isinstance(section_data, pd.DataFrame):
            # Process DataFrame
            self._update_dataframe_after_region_deletion(section_data, region_type_str, deleted_region_index, current_data)
        elif isinstance(section_data, list):
            # Process list of DataFrames
            for i, df in enumerate(section_data):
                if isinstance(df, pd.DataFrame):
                    print(f"[DEBUG] Processing DataFrame {i+1} in {region_type_str} list")
                    self._update_dataframe_after_region_deletion(df, region_type_str, deleted_region_index, current_data, list_index=i)
                else:
                    print(f"[DEBUG] Item {i+1} in {region_type_str} list is not a DataFrame: {type(df)}")
        else:
            print(f"[DEBUG] Section data for {region_type_str} is not a DataFrame or list: {type(section_data)}")

        # Update the cached extraction data
        self._cached_extraction_data = current_data
        print(f"[DEBUG] Updated cached extraction data after region deletion")

    def _update_dataframe_after_region_deletion(self, df, region_type_str, deleted_region_index, current_data, list_index=None):
        """Update a DataFrame after a region is deleted

        Args:
            df (pd.DataFrame): The DataFrame to update
            region_type_str (str): The type of region that was deleted
            deleted_region_index (int): The index of the region that was deleted
            current_data (dict): The current extraction data
            list_index (int, optional): The index of the DataFrame in a list. Defaults to None.
        """
        # Skip if the DataFrame is empty
        if df is None or df.empty:
            print(f"[DEBUG] DataFrame is empty or None, nothing to update")
            return

        # Get the title prefix for this region type
        titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
        prefix = titles.get(region_type_str, region_type_str[0].upper())
        prefix_lower = prefix.lower()
        prefix_upper = prefix.upper()

        # Check for region label column
        region_label_col = None
        for col in df.columns:
            if col == 0 or col == df.columns[0] or (isinstance(col, str) and (col.lower() == 'region_label' or col.lower() == 'region')):
                region_label_col = col
                print(f"[DEBUG] Found region label column: {col}")
                break

        if region_label_col is None:
            print(f"[DEBUG] No region label column found in DataFrame, using first column")
            if len(df.columns) > 0:
                region_label_col = df.columns[0]
            else:
                print(f"[DEBUG] DataFrame has no columns, cannot update")
                return

        # Create a new DataFrame to hold the updated data
        updated_rows = []
        updated_count = 0
        skipped_count = 0

        # Process each row in the DataFrame
        for idx, row in df.iterrows():
            # Get the current region label
            current_label = str(row[region_label_col])

            # Check if the label contains the region prefix (more flexible matching)
            if (prefix_lower in current_label.lower() or prefix_upper in current_label.upper()):
                try:
                    # Extract the region number - handle different formats
                    region_number = None

                    # Try different patterns to extract the region number
                    if '_' in current_label:
                        # Format like "H1_R1_P1"
                        parts = current_label.split('_')
                        for part in parts:
                            if (part.startswith(prefix_lower) or part.startswith(prefix_upper)) and len(part) > 1:
                                region_part = part[1:]
                                if region_part.isdigit():
                                    region_number = int(region_part)
                                    break
                    else:
                        # Format like "H1" or "h1"
                        # Use regex to find the pattern like H1, h1, etc.
                        import re
                        match = re.search(f"[{prefix_lower}{prefix_upper}](\\d+)", current_label)
                        if match:
                            region_number = int(match.group(1))

                    # If we couldn't extract a region number, try a more aggressive approach
                    if region_number is None:
                        # Extract all digits after the prefix
                        digits = ""
                        found_prefix = False
                        for i, c in enumerate(current_label):
                            if not found_prefix and (c.lower() == prefix_lower or c.upper() == prefix_upper):
                                found_prefix = True
                                continue
                            if found_prefix and c.isdigit():
                                digits += c
                            elif found_prefix and digits:
                                break

                        if digits:
                            region_number = int(digits)

                    if region_number is not None:
                        # If this row is from the deleted region, skip it
                        if region_number == deleted_region_index + 1:
                            print(f"[DEBUG] Skipping row with label {current_label} from deleted region")
                            skipped_count += 1
                            continue

                        # If this row is from a region with a higher index than the deleted one,
                        # decrement the region number
                        if region_number > deleted_region_index + 1:
                            # Create new label with decremented region number
                            # Handle different label formats properly
                            if '_' in current_label:
                                # Format like "H1_R1_P1" - preserve row and page parts
                                parts = current_label.split('_')
                                # Update only the first part (region identifier)
                                if parts[0].startswith(prefix_upper):
                                    parts[0] = f"{prefix_upper}{region_number - 1}"
                                elif parts[0].startswith(prefix_lower):
                                    parts[0] = f"{prefix_lower}{region_number - 1}"
                                # Reconstruct the label with all parts preserved
                                new_label = '_'.join(parts)
                            else:
                                # Simple format like "H1" without row or page info
                                new_label = current_label
                                # Replace the region number in the label
                                if prefix_upper in current_label:
                                    old_str = f"{prefix_upper}{region_number}"
                                    new_str = f"{prefix_upper}{region_number - 1}"
                                    new_label = new_label.replace(old_str, new_str)
                                elif prefix_lower in current_label:
                                    old_str = f"{prefix_lower}{region_number}"
                                    new_str = f"{prefix_lower}{region_number - 1}"
                                    new_label = new_label.replace(old_str, new_str)

                            # Update the region label in this row
                            row_copy = row.copy()
                            row_copy[region_label_col] = new_label
                            updated_rows.append(row_copy)
                            updated_count += 1
                            print(f"[DEBUG] Updated region label from {current_label} to {new_label}")

                            # Add more detailed debug info for multi-page labels
                            if '_P' in current_label and '_P' in new_label:
                                print(f"[DEBUG] Multi-page label updated: preserved page info from {current_label.split('_P')[1]} in {new_label}")
                            continue
                except Exception as e:
                    print(f"[DEBUG] Error parsing region label {current_label}: {str(e)}")
                    import traceback
                    traceback.print_exc()

            # Keep the row as is if we didn't skip it or update it
            updated_rows.append(row)

        # Create a new DataFrame with the updated rows
        if updated_rows:
            updated_df = pd.DataFrame(updated_rows)

            # Ensure the column order matches the original DataFrame
            updated_df = updated_df[df.columns]

            # Update the section data in the current data
            if list_index is not None:
                # Update the DataFrame in the list
                current_data[region_type_str][list_index] = updated_df
                print(f"[DEBUG] Updated DataFrame {list_index+1} in {region_type_str} list with {len(updated_df)} rows (updated: {updated_count}, skipped: {skipped_count})")
            else:
                # Update the DataFrame directly
                current_data[region_type_str] = updated_df
                print(f"[DEBUG] Updated {region_type_str} data with {len(updated_df)} rows after region deletion (updated: {updated_count}, skipped: {skipped_count})")
        else:
            # If no rows remain, set to an empty DataFrame
            if list_index is not None:
                # Update the DataFrame in the list
                current_data[region_type_str][list_index] = pd.DataFrame(columns=df.columns)
                print(f"[DEBUG] No rows remain in DataFrame {list_index+1} in {region_type_str} list")
            else:
                # Update the DataFrame directly
                current_data[region_type_str] = pd.DataFrame(columns=df.columns)
                print(f"[DEBUG] No rows remain in {region_type_str} data after region deletion")

    def _update_region_labels(self, region_type):
        """Update labels for all regions of a specific type after deletion

        Args:
            region_type (str or RegionType): The type of region to update labels for
        """
        # Convert region_type to string if it's an enum
        region_type_str = region_type.value if hasattr(region_type, 'value') else region_type

        # Get the title prefix for this region type
        titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
        prefix = titles.get(region_type_str, region_type_str[0].upper())

        if self.multi_page_mode:
            if self.current_page_index in self.page_regions and region_type_str in self.page_regions[self.current_page_index]:
                regions = self.page_regions[self.current_page_index][region_type_str]
                # Update labels for all regions
                for i, region_item in enumerate(regions):
                    if isinstance(region_item, dict) and 'rect' in region_item:
                        # Update the label to match the new index
                        region_item['label'] = f"{prefix}{i+1}"
                        print(f"[DEBUG] Updated label for {region_type_str} region at index {i} to {region_item['label']}")
        else:
            if region_type_str in self.regions:
                regions = self.regions[region_type_str]
                # Update labels for all regions
                for i, region_item in enumerate(regions):
                    if isinstance(region_item, dict) and 'rect' in region_item:
                        # Update the label to match the new index
                        region_item['label'] = f"{prefix}{i+1}"
                        print(f"[DEBUG] Updated label for {region_type_str} region at index {i} to {region_item['label']}")

    def delete_region(self, region_type, region_index):
        """Delete a specific region by type and index

        This method can be triggered by:
        1. Clicking on the X icon in the top-right corner of a region
        2. Pressing the Delete key while hovering over a region
        """
        if not self.pdf_document:
            return

        print(f"[DEBUG] Deleting region: {region_type} at index {region_index}")

        # Save the current state before deleting
        self.save_undo_state()

        # Convert region_type to RegionType enum if it's a string
        region_type_enum = RegionType(region_type) if isinstance(region_type, str) else region_type
        region_type_str = region_type_enum.value if hasattr(region_type_enum, 'value') else region_type

        # Delete the region based on multi-page mode
        if self.multi_page_mode:
            if self.current_page_index in self.page_regions:
                if region_type_str in self.page_regions[self.current_page_index]:
                    regions = self.page_regions[self.current_page_index][region_type_str]
                    if 0 <= region_index < len(regions):
                        # Remove the region
                        regions.pop(region_index)
                        print(f"[DEBUG] Deleted {region_type_str} region at index {region_index} on page {self.current_page_index + 1}")

                        # Also delete any columns associated with this region
                        if self.current_page_index in self.page_column_lines:
                            # Get the column lines for this region type
                            if region_type_enum in self.page_column_lines[self.current_page_index]:
                                # Find columns associated with this region index
                                columns_to_delete = []
                                for i, column in enumerate(self.page_column_lines[self.current_page_index][region_type_enum]):
                                    # Check if column has region index (in position 2 of tuple)
                                    if isinstance(column, tuple) and len(column) > 2 and column[2] == region_index:
                                        columns_to_delete.append(i)
                                    # For columns associated with higher region indices, decrement their index
                                    elif isinstance(column, tuple) and len(column) > 2 and column[2] > region_index:
                                        # Create new column with decremented region index
                                        new_column = (column[0], column[1], column[2] - 1)
                                        self.page_column_lines[self.current_page_index][region_type_enum][i] = new_column

                                # Delete columns in reverse order to avoid index shifting
                                for i in sorted(columns_to_delete, reverse=True):
                                    self.page_column_lines[self.current_page_index][region_type_enum].pop(i)
                                    print(f"[DEBUG] Deleted column at index {i} associated with region {region_index}")

                        # Update labels for remaining regions
                        self._update_region_labels(region_type_str)
        else:
            if region_type_str in self.regions:
                regions = self.regions[region_type_str]
                if 0 <= region_index < len(regions):
                    # Remove the region
                    regions.pop(region_index)
                    print(f"[DEBUG] Deleted {region_type_str} region at index {region_index}")

                    # Also delete any columns associated with this region
                    if region_type_enum in self.column_lines:
                        # Find columns associated with this region index
                        columns_to_delete = []
                        for i, column in enumerate(self.column_lines[region_type_enum]):
                            # Check if column has region index (in position 2 of tuple)
                            if isinstance(column, tuple) and len(column) > 2 and column[2] == region_index:
                                columns_to_delete.append(i)
                            # For columns associated with higher region indices, decrement their index
                            elif isinstance(column, tuple) and len(column) > 2 and column[2] > region_index:
                                # Create new column with decremented region index
                                new_column = (column[0], column[1], column[2] - 1)
                                self.column_lines[region_type_enum][i] = new_column

                        # Delete columns in reverse order to avoid index shifting
                        for i in sorted(columns_to_delete, reverse=True):
                            self.column_lines[region_type_enum].pop(i)
                            print(f"[DEBUG] Deleted column at index {i} associated with region {region_index}")

                    # Update labels for remaining regions
                    self._update_region_labels(region_type_str)

        # Update table_areas to remove any areas associated with the deleted region
        self._update_table_areas_after_region_deletion(region_type, region_index)

        # Get the current extraction data before updating
        current_data = self._get_current_json_data()
        print(f"[DEBUG] Current data before region deletion: {list(current_data.keys())}")

        # Update extraction results to reflect the changes in region labels
        self._update_extraction_results_after_region_deletion(region_type_str, region_index)

        # The _update_extraction_results_after_region_deletion method has already updated the cached data
        # No need to get it again or manually update it
        print(f"[DEBUG] Extraction results updated after region deletion")

        # Clear the last extraction state to force a re-extraction
        self._last_extraction_state = None

        # Force a new extraction to update the display
        self.update_extraction_results(force=True)

        # If we're in multi-page mode, update the stored data for the current page
        if self.multi_page_mode and hasattr(self, '_all_pages_data') and self._all_pages_data:
            if self.current_page_index < len(self._all_pages_data) and self._all_pages_data[self.current_page_index] is not None:
                # Update the data for the current page with the new extraction results
                self._all_pages_data[self.current_page_index] = self._cached_extraction_data.copy()
                print(f"[DEBUG] Updated stored extraction data for page {self.current_page_index + 1} after region deletion")
        self._last_extraction_state = None

        # Update the JSON tree
        self.update_json_tree(self._cached_extraction_data)

        # Update the display
        self.pdf_label.update()

    def _update_column_labels(self, column_type):
        """Update column indices after a column is deleted

        Args:
            column_type (RegionType or str): The type of column to update
        """
        # Convert column_type to RegionType enum if it's a string
        column_type_enum = RegionType(column_type) if isinstance(column_type, str) else column_type

        # Note: Column labels are automatically generated in the paintEvent method
        # based on the index of the column in the list, so we don't need to update them here.
        # This method is included for consistency with region label updates.
        print(f"[DEBUG] Column labels will be updated automatically in paintEvent")

    def delete_column(self, column_type, column_index):
        """Delete a specific column by type and index

        This method can be triggered by:
        1. Clicking on the X icon above a column line
        2. Pressing the Delete key while hovering over a column
        """
        if not self.pdf_document:
            return

        print(f"[DEBUG] Deleting column: {column_type} at index {column_index}")

        # Save the current state before deleting
        self.save_undo_state()

        # Convert column_type to RegionType enum if it's a string
        column_type_enum = RegionType(column_type) if isinstance(column_type, str) else column_type

        print(f"[DEBUG] Column type: {column_type}, converted to: {column_type_enum}")

        # Delete the column based on multi-page mode
        if self.multi_page_mode:
            if self.current_page_index in self.page_column_lines:
                if column_type_enum in self.page_column_lines[self.current_page_index]:
                    columns = self.page_column_lines[self.current_page_index][column_type_enum]
                    if 0 <= column_index < len(columns):
                        # Remove the column
                        columns.pop(column_index)
                        print(f"[DEBUG] Deleted column at index {column_index} for {column_type_enum} on page {self.current_page_index + 1}")

                        # Update column labels (not needed for columns as they're auto-numbered in paintEvent)
                        self._update_column_labels(column_type_enum)
                else:
                    print(f"[DEBUG] Column type {column_type_enum} not found in page_column_lines for page {self.current_page_index}")
                    print(f"[DEBUG] Available column types: {list(self.page_column_lines[self.current_page_index].keys())}")
        else:
            if column_type_enum in self.column_lines:
                columns = self.column_lines[column_type_enum]
                if 0 <= column_index < len(columns):
                    # Remove the column
                    columns.pop(column_index)
                    print(f"[DEBUG] Deleted column at index {column_index} for {column_type_enum}")

                    # Update column labels (not needed for columns as they're auto-numbered in paintEvent)
                    self._update_column_labels(column_type_enum)
            else:
                print(f"[DEBUG] Column type {column_type_enum} not found in column_lines")
                print(f"[DEBUG] Available column types: {list(self.column_lines.keys())}")

        # Clear cached extraction data for this page
        self._cached_extraction_data = {
            'header': [],
            'items': [],
            'summary': []
        }
        self._last_extraction_state = None

        # Update the JSON tree
        self.update_json_tree(self._cached_extraction_data)

        # Update the display
        self.pdf_label.update()

        # Force extraction to update results
        self.update_extraction_results(force=True)

        # If we're in multi-page mode, update the stored data for the current page
        if self.multi_page_mode and hasattr(self, '_all_pages_data') and self._all_pages_data:
            if self.current_page_index < len(self._all_pages_data) and self._all_pages_data[self.current_page_index] is not None:
                # Update the data for the current page with the new extraction results
                self._all_pages_data[self.current_page_index] = self._cached_extraction_data.copy()
                print(f"[DEBUG] Updated stored extraction data for page {self.current_page_index + 1} after column deletion")

    def save_undo_state(self):
        """Save the current state for undo"""
        # Undo feature has been removed as per user preference
        # This method is kept for compatibility with existing code
        print("[DEBUG] Undo feature is disabled, but save_undo_state was called")

        # No need to create a deep copy or append to undo_stack
        # Just return without doing anything
        return

    def initUI(self):
        """Initialize the user interface with a splitter"""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Create and add top bar
        top_bar = self._create_top_bar()
        main_layout.addWidget(top_bar)

        # Create main splitter and sections
        self._create_main_splitter()
        self._create_pdf_container()
        self._create_extraction_viewer()
        self._create_invoice2data_container()
        self._setup_splitter_layout()

        # Add splitter to main layout
        main_layout.addWidget(self.main_splitter)

    def _create_top_bar(self):
        """Create the top bar with global controls"""
        top_bar = QWidget()
        top_bar.setFixedHeight(50)
        top_bar.setStyleSheet("background-color: #f0f0f0; border-bottom: 1px solid #ddd;")
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(10, 5, 10, 5)

        # Add back button on the left
        self.global_back_btn = QPushButton("← Back")
        self.global_back_btn.clicked.connect(self.go_back.emit)
        self.global_back_btn.setStyleSheet("""
            QPushButton {
                background-color: #555555;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
                min-width: 50px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #333333;
            }
        """)
        top_bar_layout.addWidget(self.global_back_btn)

        # Add Reset Screen button next to back button
        self.reset_screen_btn = QPushButton("Reset Screen")
        self.reset_screen_btn.clicked.connect(self.reset_screen)
        self.reset_screen_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme['danger']};
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
                min-width: 50px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {self.theme['danger'] + 'E0'};
            }}
        """)
        top_bar_layout.addWidget(self.reset_screen_btn)

        # Add spacer to push buttons to the right
        top_bar_layout.addStretch(1)

        # Create container for global buttons
        global_buttons = QHBoxLayout()
        global_buttons.setSpacing(10)

        # Save template button
        self.global_save_template_btn = QPushButton("Save Template")
        self.global_save_template_btn.clicked.connect(self.save_template)
        self.global_save_template_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
                min-width: 50px;
            }
            QPushButton:hover {
                background-color: #3159C1;
            }
        """)
        global_buttons.addWidget(self.global_save_template_btn)

        # Add the global buttons to the top bar
        top_bar_layout.addLayout(global_buttons)

        return top_bar

    def _create_main_splitter(self):
        """Create the main horizontal splitter"""
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setHandleWidth(3)
        self.main_splitter.setChildrenCollapsible(True)
        self.main_splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #cccccc;
            }
            QSplitter::handle:hover {
                background-color: #3498db;
            }
        """)

        # Connect splitter moved signal
        self.main_splitter.splitterMoved.connect(self.on_splitter_moved)

        # Initialize user adjusted splitter flag
        self._user_adjusted_splitter = False
        self.pdf_section_was_hidden = False

    def _create_pdf_container(self):
        """Create the PDF container with controls and display area"""
        self.pdf_container = QWidget()
        self.pdf_container.setStyleSheet("background-color: #000000;")
        pdf_layout = QVBoxLayout(self.pdf_container)
        pdf_layout.setContentsMargins(10, 10, 10, 10)

        # Title for PDF side
        pdf_title = QLabel("Invoice Designer")
        pdf_title.setFont(QFont("Arial", 16, QFont.Bold))
        pdf_title.setAlignment(Qt.AlignCenter)
        pdf_title.setStyleSheet("background-color: #000000; color: white; padding: 8px; border-bottom: 1px solid #333;")
        pdf_layout.addWidget(pdf_title)

        # Add navigation controls
        self._create_pdf_navigation_controls(pdf_layout)

        # Add PDF controls container
        self._create_pdf_controls_container(pdf_layout)

        # Add PDF display area
        self._create_pdf_display_area(pdf_layout)

    def _create_pdf_navigation_controls(self, pdf_layout):
        """Create navigation controls for PDF"""
        top_controls_layout = QHBoxLayout()

        # Left section (empty now as back button moved to top bar)
        left_controls = QHBoxLayout()
        left_controls.setAlignment(Qt.AlignLeft)
        left_controls_container = QWidget()
        left_controls_container.setStyleSheet("background-color: #000000;")
        left_controls_container.setLayout(left_controls)
        top_controls_layout.addWidget(left_controls_container, 1)

        # Center controls (navigation buttons)
        center_controls = QHBoxLayout()

        # Page navigation buttons (initially hidden)
        self.prev_page_btn = QPushButton("← Previous")
        self.prev_page_btn.clicked.connect(self.prev_page)
        self.prev_page_btn.hide()
        self.prev_page_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #3159C1;
            }
        """)
        center_controls.addWidget(self.prev_page_btn)

        self.next_page_btn = QPushButton("Next →")
        self.next_page_btn.clicked.connect(self.next_page)
        self.next_page_btn.hide()
        self.next_page_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #3159C1;
            }
        """)
        center_controls.addWidget(self.next_page_btn)

        # Apply to remaining pages button (initially hidden)
        self.apply_to_remaining_btn = QPushButton("Apply to all")
        self.apply_to_remaining_btn.clicked.connect(self.apply_to_remaining_pages)
        self.apply_to_remaining_btn.hide()
        self.apply_to_remaining_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        center_controls.addWidget(self.apply_to_remaining_btn)

        # Add center controls to main controls layout
        center_controls.setAlignment(Qt.AlignCenter)
        center_controls_container = QWidget()
        center_controls_container.setStyleSheet("background-color: #000000;")
        center_controls_container.setLayout(center_controls)
        top_controls_layout.addWidget(center_controls_container, 1)

        # Right controls (empty now as buttons moved to top bar)
        right_controls = QHBoxLayout()
        right_controls.setAlignment(Qt.AlignRight)
        right_controls_container = QWidget()
        right_controls_container.setStyleSheet("background-color: #000000;")
        right_controls_container.setLayout(right_controls)
        top_controls_layout.addWidget(right_controls_container, 1)

        # Add the controls layout to the main layout
        pdf_layout.addLayout(top_controls_layout)

    def _create_pdf_controls_container(self, pdf_layout):
        """Create the PDF controls container with region buttons"""
        self.pdf_controls_container = QWidget()
        self.pdf_controls_container.setStyleSheet("background-color: #000000;")
        pdf_controls_layout = QVBoxLayout(self.pdf_controls_container)
        pdf_controls_layout.setContentsMargins(0, 0, 0, 0)

        # Create region buttons
        self._create_region_buttons(pdf_controls_layout)

        # Add the PDF controls container to the main PDF layout
        pdf_layout.addWidget(self.pdf_controls_container)

        # Hide PDF controls initially until a PDF is loaded
        self.pdf_controls_container.hide()

    def _create_region_buttons(self, pdf_controls_layout):
        """Create region selection buttons"""
        region_layout = QHBoxLayout()

        # Common button style
        button_style = """
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                padding: 4px 8px;
                border-radius: 4px;
                min-width: 50px;
                font-weight: bold;
                color: black;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """

        button_width = 120

        # Header region button
        self.header_btn = QPushButton("Header")
        self.header_btn.setCheckable(True)
        self.header_btn.setFixedWidth(button_width)
        self.header_btn.clicked.connect(lambda: self.set_region_type('header'))
        self.header_btn.setStyleSheet(button_style + """
            QPushButton:checked {
                background-color: #3498db;
                color: black;
                border: 1px solid #2980b9;
                font-weight: bold;
            }
        """)
        region_layout.addWidget(self.header_btn)

        # Items region button
        self.items_btn = QPushButton("Items")
        self.items_btn.setCheckable(True)
        self.items_btn.setFixedWidth(button_width)
        self.items_btn.clicked.connect(lambda: self.set_region_type('items'))
        self.items_btn.setStyleSheet(button_style + """
            QPushButton:checked {
                background-color: #2ecc71;
                color: black;
                border: 1px solid #27ae60;
                font-weight: bold;
            }
        """)
        region_layout.addWidget(self.items_btn)

        # Summary region button
        self.summary_btn = QPushButton("Summary")
        self.summary_btn.setCheckable(True)
        self.summary_btn.setFixedWidth(button_width)
        self.summary_btn.clicked.connect(lambda: self.set_region_type('summary'))
        self.summary_btn.setStyleSheet(button_style + """
            QPushButton:checked {
                background-color: #9b59b6;
                color: black;
                border: 1px solid #8e44ad;
                font-weight: bold;
            }
        """)
        region_layout.addWidget(self.summary_btn)

        # Column drawing button
        self.column_btn = QPushButton("Columns")
        self.column_btn.setCheckable(True)
        self.column_btn.setFixedWidth(button_width)
        self.column_btn.clicked.connect(self.toggle_column_drawing)
        self.column_btn.setStyleSheet(button_style + """
            QPushButton:checked {
                background-color: #f39c12;
                color: black;
                border: 1px solid #e67e22;
                font-weight: bold;
            }
        """)
        region_layout.addWidget(self.column_btn)

        # Clear Screen button
        self.clear_screen_btn = QPushButton("Clear Drawing")
        self.clear_screen_btn.setFixedWidth(button_width)
        self.clear_screen_btn.clicked.connect(self.clear_current_page)
        self.clear_screen_btn.setStyleSheet(button_style)
        region_layout.addWidget(self.clear_screen_btn)

        # Auto column mode button
        self.auto_column_btn = QPushButton("Column: Auto")
        self.auto_column_btn.setCheckable(True)
        self.auto_column_btn.setChecked(False)
        self.auto_column_btn.clicked.connect(self.toggle_auto_column_mode)
        self.auto_column_btn.setFixedWidth(120)
        self.auto_column_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                padding: 4px 8px;
                border-radius: 4px;
                color: black;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #e0e0e0;
            }}
        """)
        region_layout.addWidget(self.auto_column_btn)

        # Add region layout to the PDF controls container
        pdf_controls_layout.addLayout(region_layout)

    def _create_pdf_display_area(self, pdf_layout):
        """Create the PDF display area with scroll area and upload area"""
        # PDF display area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(False)
        self.scroll_area.setMinimumSize(QSize(900, 1200))
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # Style the scroll area
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                background-color: #000000;
                border: 1px solid #000000;
            }
            QScrollArea > QWidget > QWidget {
                background-color: #000000;
                padding-bottom: 100px;
                margin-bottom: 100px;
            }
            QScrollBar:vertical {
                width: 16px;
            }
            QScrollBar:horizontal {
                height: 16px;
            }
        """)

        # Create PDF label for displaying the PDF
        self.pdf_label = PDFLabel(self)
        self.pdf_label.setAlignment(Qt.AlignCenter)
        self.pdf_label.setStyleSheet("QLabel { background-color: #000000; }")

        # Connect zoom signal from PDFLabel
        try:
            self.pdf_label.zoom_changed.connect(self.update_zoom_label)
            print(f"[DEBUG] Successfully connected zoom_changed signal")
        except Exception as e:
            print(f"[DEBUG] Error connecting zoom_changed signal: {str(e)}")

        # Connect to scroll area's viewport events
        self.scroll_area.viewport().installEventFilter(self)

        # Set minimum size for PDF label
        self.pdf_label.setMinimumSize(QSize(620, 870))

        # Create upload area
        self._create_upload_area(pdf_layout)

        # Add PDF label to scroll area
        self.scroll_area.setWidget(self.pdf_label)
        self.pdf_label.hide()
        pdf_layout.addWidget(self.scroll_area)

        # Create floating zoom controls
        self.create_floating_zoom_controls()

    def _create_upload_area(self, pdf_layout):
        """Create the upload area for drag and drop"""
        self.upload_area = ClickableLabel("Drag & Drop PDF or Click to Upload")
        self.upload_area.setAlignment(Qt.AlignCenter)
        self.upload_area.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 5px;
                padding: 20px;
                background-color: #000000;
                color: #ffffff;
                font-size: 16px;
            }
            QLabel:hover {
                background-color: #222222;
                border-color: #999;
            }
        """)
        self.upload_area.setMinimumHeight(300)
        self.upload_area.clicked.connect(self.load_pdf)

        # Set up drag and drop
        self.upload_area.setAcceptDrops(True)
        self.setAcceptDrops(True)

        # Add upload area to PDF layout
        pdf_layout.addWidget(self.upload_area)

    def _create_extraction_viewer(self):
        """Create the extraction results viewer section"""
        self.extraction_viewer = QWidget()
        extraction_layout = QVBoxLayout(self.extraction_viewer)
        extraction_layout.setContentsMargins(10, 10, 10, 10)

        # Add title
        extraction_title = QLabel("Extraction Results")
        extraction_title.setFont(QFont("Arial", 12, QFont.Bold))
        extraction_title.setAlignment(Qt.AlignCenter)
        extraction_title.setStyleSheet("background-color: #000000; color: white; padding: 8px; border-bottom: 1px solid #333;")
        extraction_layout.addWidget(extraction_title)

        # Create tree container
        self._create_tree_container(extraction_layout)

        # Add extraction controls
        self._create_extraction_controls(extraction_layout)

    def _create_tree_container(self, extraction_layout):
        """Create the tree container with JSON tree and controls"""
        self.tree_container = QWidget()
        tree_layout = QVBoxLayout(self.tree_container)

        # Initialize the JSON tree view
        self.json_tree = QTreeWidget()
        self.json_tree.setHeaderLabels(["Field", "Value"])
        self.json_tree.setAlternatingRowColors(False)
        self.json_tree.setColumnWidth(0, 250)
        self.json_tree.setColumnWidth(1, 350)

        # Initialize with placeholder message
        self.json_tree.clear()
        placeholder_item = QTreeWidgetItem(["No data", "No data extracted yet. Draw regions on the PDF to see results."])
        self.json_tree.addTopLevelItem(placeholder_item)
        tree_layout.addWidget(self.json_tree)

        # Add control buttons
        tree_controls = QHBoxLayout()

        copy_all_btn = QPushButton("Copy All to Clipboard")
        copy_all_btn.clicked.connect(self.copy_all_data_to_clipboard)
        copy_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
        """)

        # Add toggle button for JSON Designer
        self.toggle_json_designer_btn = QPushButton("JSON Designer")
        self.toggle_json_designer_btn.setCheckable(True)
        self.toggle_json_designer_btn.setChecked(False)
        self.toggle_json_designer_btn.clicked.connect(self.toggle_json_designer)
        self.toggle_json_designer_btn.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3159C1;
            }
            QPushButton:checked {
                background-color: #3159C1;
            }
        """)

        tree_controls.addWidget(copy_all_btn)
        tree_controls.addWidget(self.toggle_json_designer_btn)
        tree_layout.addLayout(tree_controls)

        extraction_layout.addWidget(self.tree_container)

    def _create_extraction_controls(self, extraction_layout):
        """Create extraction control buttons"""
        # Add extraction method selection
        method_layout = QHBoxLayout()
        method_label = QLabel("Extraction Method:")
        method_label.setStyleSheet("color: black; font-weight: bold;")

        self.extraction_method_combo = QComboBox()
        self.extraction_method_combo.addItems([
            "pypdf_table_extraction",
            "pdftotext",
            "tesseract_ocr"
        ])
        self.extraction_method_combo.setCurrentText("pypdf_table_extraction")
        self.extraction_method_combo.currentTextChanged.connect(self.on_extraction_method_changed)
        self.extraction_method_combo.setStyleSheet("""
            QComboBox {
                background-color: white;
                border: 1px solid #ddd;
                padding: 4px 8px;
                border-radius: 4px;
                min-width: 150px;
                color: black;
            }
            QComboBox:hover {
                border-color: #4169E1;
            }
        """)

        method_layout.addWidget(method_label)
        method_layout.addWidget(self.extraction_method_combo)
        method_layout.addStretch()
        extraction_layout.addLayout(method_layout)

        extraction_controls = QHBoxLayout()

        self.retry_btn = QPushButton("Retry Extraction")
        self.retry_btn.clicked.connect(self.retry_extraction)
        self.retry_btn.setEnabled(True)
        self.retry_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 120px;
                color: black;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)

        self.adjust_params_btn = QPushButton("Adjust Parameters")
        self.adjust_params_btn.clicked.connect(self.show_param_dialog)
        self.adjust_params_btn.setEnabled(True)
        self.adjust_params_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 120px;
                color: black;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)

        extraction_controls.addWidget(self.retry_btn)
        extraction_controls.addWidget(self.adjust_params_btn)
        extraction_layout.addLayout(extraction_controls)

    def _create_invoice2data_container(self):
        """Create the invoice2data container section"""
        self.invoice2data_container = QWidget()
        invoice2data_layout = QVBoxLayout(self.invoice2data_container)
        invoice2data_layout.setContentsMargins(10, 10, 10, 10)

        # Add title
        invoice2data_title = QLabel("JSON Designer")
        invoice2data_title.setFont(QFont("Arial", 12, QFont.Bold))
        invoice2data_title.setAlignment(Qt.AlignCenter)
        invoice2data_title.setStyleSheet("background-color: #000000; color: white; padding: 8px; border-bottom: 1px solid #333;")
        invoice2data_layout.addWidget(invoice2data_title)

        # Create the invoice2data editor
        self.invoice2data_editor = self.create_invoice2data_editor()
        invoice2data_layout.addWidget(self.invoice2data_editor)

    def _setup_splitter_layout(self):
        """Setup the splitter layout with all sections"""
        # Add all three sections to the main splitter
        self.main_splitter.addWidget(self.pdf_container)
        self.main_splitter.addWidget(self.extraction_viewer)
        self.main_splitter.addWidget(self.invoice2data_container)

        # Set initial sizes (with JSON designer section hidden)
        total_width = 1500
        self.main_splitter.setSizes([int(total_width * 0.5), int(total_width * 0.5), 0])

        # Hide the JSON designer section by default
        self.invoice2data_container.hide()

    def toggle_json_designer(self):
        """Toggle the visibility of the JSON Designer tabs in the extraction results section"""
        # Get the current state of the toggle button
        is_visible = self.toggle_json_designer_btn.isChecked()

        # Get current sizes
        sizes = self.main_splitter.sizes()

        if is_visible:
            # Make sure we have a template loaded
            if not hasattr(self, 'invoice2data_template') or self.invoice2data_template is None:
                self.initialize_invoice2data_template()
                print(f"[DEBUG] Initialized invoice2data template in toggle_json_designer")

            # Populate the form with the template values
            self.populate_form_from_template()
            print(f"[DEBUG] Populated form from template in toggle_json_designer")

            # Ensure the header fields table is populated
            if hasattr(self, 'header_fields_table') and self.header_fields_table.rowCount() == 0 and 'fields' in self.invoice2data_template:
                fields = self.invoice2data_template['fields']

                # Add fields from template
                for field_name, field_data in fields.items():
                    row = self.header_fields_table.rowCount()
                    self.header_fields_table.insertRow(row)

                    # Set field name
                    self.header_fields_table.setItem(row, 0, QTableWidgetItem(field_name))

                    # Set regex pattern
                    if isinstance(field_data, dict) and 'regex' in field_data:
                        regex_pattern = field_data['regex']
                    else:
                        # Simple field format (just a regex pattern)
                        regex_pattern = str(field_data)
                    self.header_fields_table.setItem(row, 1, QTableWidgetItem(regex_pattern))

                    # Set field type
                    type_combo = QComboBox()
                    type_combo.addItems(["string", "date", "float", "int"])
                    if isinstance(field_data, dict) and 'type' in field_data:
                        field_type = field_data['type']
                        if field_type in ["string", "date", "float", "int"]:
                            type_combo.setCurrentText(field_type)
                    self.header_fields_table.setCellWidget(row, 2, type_combo)

                print(f"[DEBUG] Directly populated header fields table with {self.header_fields_table.rowCount()} fields in toggle_json_designer")

            # Show the appropriate JSON Designer tabs based on the current bottom tab
            if hasattr(self, 'bottom_tabs'):
                current_tab_index = self.bottom_tabs.currentIndex()

                # Show only the editors for the currently selected tab
                if current_tab_index == 0:  # Header
                    if hasattr(self, 'header_fields_editor'):
                        # Make sure the Fields tab is selected in the header tab widget
                        if hasattr(self, 'header_tab_widget'):
                            fields_tab_index = self.header_tab_widget.indexOf(self.header_fields_editor)
                            if fields_tab_index >= 0:
                                self.header_tab_widget.setCurrentIndex(fields_tab_index)
                                print(f"[DEBUG] Activated Fields tab in header tab widget")
                elif current_tab_index == 1:  # Items
                    if hasattr(self, 'items_tab_widget'):
                        # Show both Tables and Lines tabs, but select Tables by default
                        tables_tab_index = self.items_tab_widget.indexOf(self.items_tables_editor)
                        if tables_tab_index >= 0:
                            self.items_tab_widget.setCurrentIndex(tables_tab_index)
                            print(f"[DEBUG] Activated Tables tab in items tab widget")
                elif current_tab_index == 2:  # Summary
                    if hasattr(self, 'summary_tab_widget'):
                        # Show both Fields and Tax Lines tabs, but select Fields by default
                        fields_tab_index = self.summary_tab_widget.indexOf(self.summary_fields_editor)
                        if fields_tab_index >= 0:
                            self.summary_tab_widget.setCurrentIndex(fields_tab_index)
                            print(f"[DEBUG] Activated Fields tab in summary tab widget")

            # Show the JSON Designer container (for backward compatibility)
            if hasattr(self, 'invoice2data_container'):
                self.invoice2data_container.show()

                # Adjust splitter sizes to show JSON Designer with a 4:4:2 ratio
                total_width = sum(sizes)
                self.main_splitter.setSizes([
                    int(total_width * 0.4),
                    int(total_width * 0.4),
                    int(total_width * 0.2)
                ])

            print(f"[DEBUG] JSON Designer shown")
        else:
            # Hide all JSON Designer tabs
            if hasattr(self, 'header_fields_editor'):
                self.header_fields_editor.hide()
            if hasattr(self, 'items_tables_editor'):
                self.items_tables_editor.hide()
            if hasattr(self, 'items_lines_editor'):
                self.items_lines_editor.hide()
            if hasattr(self, 'summary_fields_editor'):
                self.summary_fields_editor.hide()
            if hasattr(self, 'tax_lines_editor'):
                self.tax_lines_editor.hide()

            # Hide the JSON Designer container (for backward compatibility)
            if hasattr(self, 'invoice2data_container'):
                self.invoice2data_container.hide()

                # Adjust splitter sizes to hide JSON Designer with a 1:1:0 ratio
                total_width = sum(sizes)
                self.main_splitter.setSizes([
                    int(total_width * 0.5),
                    int(total_width * 0.5),
                    0
                ])

            print(f"[DEBUG] JSON Designer hidden")

    def on_splitter_moved(self, pos, index):
        """Handle splitter moved event to maintain section visibility"""
        # Get current sizes
        sizes = self.main_splitter.sizes()
        print(f"[DEBUG] Splitter sections: {sizes}")

        # Set the user adjusted flag to true
        self._user_adjusted_splitter = True

        # Update the toggle button state based on JSON Designer visibility
        if sizes[2] > 0:
            self.toggle_json_designer_btn.setChecked(True)
        else:
            self.toggle_json_designer_btn.setChecked(False)

    def set_region_type(self, region_type):
        """Set the current region type for drawing"""
        # Define button styles based on theme
        default_style = f"""
            QPushButton {{
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                padding: 4px 8px;
                border-radius: 4px;
                color: black;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #e0e0e0;
            }}
        """

        # Uncheck all buttons first
        self.header_btn.setChecked(False)
        self.items_btn.setChecked(False)
        self.summary_btn.setChecked(False)

        # Uncheck column button and update its style
        self.column_btn.setChecked(False)
        self.column_btn.setStyleSheet(default_style)

        # Set the current region type
        self.current_region_type = region_type
        self.drawing_column = False

        # Check the appropriate button and update styles
        if region_type == 'header':
            self.header_btn.setChecked(True)
            # Make the header tab active in both top and bottom sections
            if hasattr(self, 'top_tabs'):
                # Find the index of the header tab in top section
                header_tab_index = self.top_tabs.indexOf(self.header_raw_text_tab)
                if header_tab_index >= 0:
                    self.top_tabs.setCurrentIndex(header_tab_index)
                    print(f"[DEBUG] Activated header tab in top section")
                    # The on_top_tab_changed method will handle synchronizing the bottom tabs

        elif region_type == 'items':
            self.items_btn.setChecked(True)
            # Make the items tab active in both top and bottom sections
            if hasattr(self, 'top_tabs'):
                # Find the index of the items tab in top section
                items_tab_index = self.top_tabs.indexOf(self.items_raw_text_tab)
                if items_tab_index >= 0:
                    self.top_tabs.setCurrentIndex(items_tab_index)
                    print(f"[DEBUG] Activated items tab in top section")
                    # The on_top_tab_changed method will handle synchronizing the bottom tabs

        elif region_type == 'summary':
            self.summary_btn.setChecked(True)
            # Make the summary tab active in both top and bottom sections
            if hasattr(self, 'top_tabs'):
                # Find the index of the summary tab in top section
                summary_tab_index = self.top_tabs.indexOf(self.summary_raw_text_tab)
                if summary_tab_index >= 0:
                    self.top_tabs.setCurrentIndex(summary_tab_index)
                    print(f"[DEBUG] Activated summary tab in top section")
                    # The on_top_tab_changed method will handle synchronizing the bottom tabs

        # Set cursor for drawing
        if self.pdf_label:
            self.pdf_label.setCursor(Qt.CrossCursor)

        # Update the display
        self.pdf_label.update()

    def auto_detect_regions(self, pos):
        """Auto-detect if the mouse is hovering over a region and switch to column drawing mode"""
        if not self.auto_column_mode or not self.pdf_document or self.pdf_label.drawing:
            return

        # Check if we're already in column drawing mode
        if self.drawing_column:
            return

        # Limit the frequency of cursor updates to avoid flickering
        import time
        current_time = time.time()
        if current_time - self.last_cursor_update < 0.1:  # 100ms
            return
        self.last_cursor_update = current_time

        # Check each region to see if the mouse is inside
        found_region = False
        for region_type, rects in self.regions.items():
            for i, region_item in enumerate(rects):
                # Check if the region is stored as a dict with 'rect' and 'label'
                if isinstance(region_item, dict) and 'rect' in region_item:
                    rect = region_item['rect']
                else:
                    # Backward compatibility for old format
                    rect = region_item

                if rect.contains(pos):
                    # We found a region under the cursor
                    found_region = True
                    self.hover_region_type = region_type
                    self.hover_rect_index = i

                    # Change cursor to indicate column drawing is available
                    if self.pdf_label:
                        self.pdf_label.setCursor(Qt.SplitHCursor)

                    # Store the active region for column drawing
                    self.active_region_type = region_type
                    self.active_rect_index = i

                    # Make sure the current_rect is valid
                    if region_type in self.regions and 0 <= i < len(self.regions[region_type]):
                        self.current_rect = rect
                    else:
                        print(f"[DEBUG] Invalid region index: {region_type} {i}")
                        continue

                    # Set current region type for column drawing
                    self.current_region_type = region_type

                    # Enable column drawing mode temporarily
                    self.drawing_column = True

                    print(f"[DEBUG] Auto-detected region: {region_type} {i}")
                    return

        # If we didn't find a region, reset to normal cursor
        if not found_region:
            self.hover_region_type = None
            self.hover_rect_index = None

            # Reset column drawing mode
            self.drawing_column = False

            # Reset cursor
            if self.pdf_label:
                self.pdf_label.setCursor(Qt.ArrowCursor)

    def toggle_auto_column_mode(self):
        """Toggle automatic column drawing mode"""
        # Toggle auto column mode
        self.auto_column_mode = self.auto_column_btn.isChecked()

        # Update button text and style
        if self.auto_column_mode:
            self.auto_column_btn.setText("Column: Auto")
            self.auto_column_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {self.theme['secondary']};
                    color: black;
                    padding: 4px 8px;
                    border-radius: 4px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: #218838;
                }}
            """)
            print(f"[DEBUG] Auto column mode enabled")
        else:
            self.auto_column_btn.setText("Column: Auto")
            self.auto_column_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: #f0f0f0;
                    border: 1px solid #ddd;
                    padding: 4px 8px;
                    border-radius: 4px;
                    color: black;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: #e0e0e0;
                }}
            """)
            print(f"[DEBUG] Auto column mode disabled")

            # Reset cursor if we're not in column drawing mode
            if not self.drawing_column and self.pdf_label:
                self.pdf_label.setCursor(Qt.ArrowCursor)

    def toggle_column_drawing(self):
        """Toggle column drawing mode"""
        # Toggle column drawing mode
        self.drawing_column = self.column_btn.isChecked()

        # Define button styles based on theme
        default_style = f"""
            QPushButton {{
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                padding: 4px 8px;
                border-radius: 4px;
                color: black;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #e0e0e0;
            }}
        """

        selected_style = f"""
            QPushButton {{
                background-color: {self.theme['primary']};
                color: black;
                border: 1px solid {self.theme['primary_dark']};
                padding: 4px 8px;
                border-radius: 4px;
                font-weight: bold;
            }}
        """

        # Update button styles and state based on column drawing mode
        if self.drawing_column:
            # If column mode is active, uncheck all region buttons
            self.header_btn.setChecked(False)
            self.items_btn.setChecked(False)
            self.summary_btn.setChecked(False)

            # Apply selected style to column button
            self.column_btn.setStyleSheet(selected_style)

            # Default to 'items' if no region type is selected
            if not self.current_region_type:
                self.current_region_type = 'items'
                print(f"[DEBUG] Setting default region type to 'items' for column drawing")

            # Set cursor for column drawing
            if self.pdf_label:
                self.pdf_label.setCursor(Qt.SplitHCursor)
                print(f"[DEBUG] Column drawing mode enabled for region type: {self.current_region_type}")
        else:
            # If column mode is inactive, reset current region type
            self.current_region_type = None

            # Apply default style to column button
            self.column_btn.setStyleSheet(default_style)

            # Reset cursor
            if self.pdf_label:
                self.pdf_label.setCursor(Qt.ArrowCursor)

            print(f"[DEBUG] Column drawing mode disabled")

        # Update the display
        self.pdf_label.update()

    def handle_mouse_press(self, pos):
        """Handle mouse press event from PDFLabel"""
        # Reset flags that prevent duplication
        if hasattr(self, '_in_specific_region_extraction'):
            self._in_specific_region_extraction = False
            print(f"[DEBUG] Reset _in_specific_region_extraction flag in mouse press event")

        if hasattr(self, '_skip_extraction_update'):
            self._skip_extraction_update = False
            print(f"[DEBUG] Reset _skip_extraction_update flag in mouse press event")

        # Check if we're in auto column mode and the mouse is over a region
        if self.auto_column_mode and self.hover_region_type is not None and self.hover_rect_index is not None:
            # Verify that the hover_region_type and hover_rect_index are valid
            if (self.hover_region_type in self.regions and
                0 <= self.hover_rect_index < len(self.regions[self.hover_region_type])):
                # We're in auto column mode and hovering over a valid region
                # Start drawing a column line
                self.pdf_label.drawing = True
                self.pdf_label.start_pos = pos
                self.pdf_label.current_pos = pos
                self.pdf_label.drawing_mode = 'column'

                # Use the hover region as the active region
                self.active_region_type = self.hover_region_type
                self.active_rect_index = self.hover_rect_index
                self.pdf_label.current_region_type = self.hover_region_type

                # Get the active rectangle
                self.current_rect = self.regions[self.hover_region_type][self.hover_rect_index]

                print(f"[DEBUG] Auto column mode: Drawing column in {self.hover_region_type} region {self.hover_rect_index}")

                # Update display
                self.pdf_label.update()
                return
            else:
                # Invalid hover region or index, reset hover state
                print(f"[DEBUG] Invalid hover region or index: {self.hover_region_type}, {self.hover_rect_index}")
                self.hover_region_type = None
                self.hover_rect_index = None

        # Normal drawing mode
        if self.current_region_type or self.drawing_column:
            # Start drawing
            self.pdf_label.drawing = True
            self.pdf_label.start_pos = pos
            self.pdf_label.current_pos = pos

            # Set drawing mode and region type
            if self.drawing_column:
                self.pdf_label.drawing_mode = 'column'
                # Find which region the column is being drawn in
                self.active_region_type = None
                self.active_rect_index = None

                # Check each region to see if the click is inside
                for region_type, rects in self.regions.items():
                    for i, region_item in enumerate(rects):
                        # Use standardized coordinate system - single format everywhere
                        from standardized_coordinates import StandardRegion

                        # Enforce StandardRegion format - NO backward compatibility
                        if not isinstance(region_item, StandardRegion):
                            print(f"[ERROR] Invalid region type in {region_type}[{i}]: expected StandardRegion, got {type(region_item)}")
                            continue

                        rect = region_item.rect  # Get the QRect from StandardRegion

                        if rect.contains(pos):
                            self.active_region_type = region_type
                            self.active_rect_index = i
                            self.pdf_label.current_region_type = region_type
                            print(f"[DEBUG] Drawing column in {region_type} region {i}")

                            # Store the active rectangle for drawing preview
                            self.current_rect = rect
                            break
                    if self.active_region_type:
                        break

                # Change cursor to indicate active drawing
                self.pdf_label.setCursor(Qt.SplitHCursor)
            else:
                self.pdf_label.drawing_mode = 'region'
                self.pdf_label.current_region_type = self.current_region_type

                # Change cursor to indicate active drawing
                self.pdf_label.setCursor(Qt.CrossCursor)

            # Update display
            self.pdf_label.update()

    def handle_mouse_move(self, pos):
        """Handle mouse move event from PDFLabel"""
        if self.pdf_label.drawing:
            # Update current position
            self.pdf_label.current_pos = pos

            # If drawing column lines, constrain to the active rectangle
            if self.pdf_label.drawing_mode == 'column' and hasattr(self, 'current_rect') and self.current_rect:
                # Keep x position within the rectangle boundaries
                rect = self.current_rect
                if rect:
                    # Constrain x position to the rectangle
                    x_pos = max(rect.left(), min(pos.x(), rect.right()))
                    self.pdf_label.current_pos = QPoint(x_pos, pos.y())

                    print(f"[DEBUG] Column drawing at x={x_pos} in rectangle {rect.x()},{rect.y()},{rect.width()},{rect.height()}")

            # Update display
            self.pdf_label.update()

    def handle_mouse_release(self, pos):
        """Handle mouse release event from PDFLabel"""
        if self.pdf_label.drawing:
            # Finish drawing
            self.pdf_label.drawing = False

            # If we're in auto column mode and just finished drawing a column,
            # reset the drawing_column flag to allow auto-detection to work again
            if self.auto_column_mode and self.pdf_label.drawing_mode == 'column':
                # Reset drawing column mode but keep the active region information
                was_auto_column = (self.hover_region_type is not None and self.hover_rect_index is not None)
                if was_auto_column:
                    # Verify that the hover_region_type and hover_rect_index are valid
                    if (self.hover_region_type in self.regions and
                        0 <= self.hover_rect_index < len(self.regions[self.hover_region_type])):
                        print(f"[DEBUG] Auto column mode: Finished drawing column, resetting drawing_column flag")
                        self.drawing_column = False
                    else:
                        # Invalid hover region or index, reset hover state
                        print(f"[DEBUG] Invalid hover region or index in release: {self.hover_region_type}, {self.hover_rect_index}")
                        self.hover_region_type = None
                        self.hover_rect_index = None

            # Create rectangle or column line based on drawing mode
            if self.pdf_label.drawing_mode == 'region' and self.pdf_label.start_pos and pos:
                # Create rectangle
                rect = QRect(self.pdf_label.start_pos, pos).normalized()

                # Add to regions if valid
                if rect.width() > 10 and rect.height() > 10:
                    # Create region label based on region type and index
                    titles = {'header': 'H', 'items': 'I', 'summary': 'S'}

                    if self.multi_page_mode:
                        # For multi-page mode, add to page_regions
                        if self.current_page_index not in self.page_regions:
                            self.page_regions[self.current_page_index] = {'header': [], 'items': [], 'summary': []}
                            # Also initialize the regions dictionary to match
                            self.regions = self.page_regions[self.current_page_index]

                        # Make sure the current_region_type exists in the regions dictionary
                        if self.current_region_type not in self.regions:
                            self.regions[self.current_region_type] = []
                            self.page_regions[self.current_page_index][self.current_region_type] = []
                            print(f"[DEBUG] Initialized empty {self.current_region_type} list for page {self.current_page_index + 1}")

                        region_index = len(self.page_regions[self.current_page_index][self.current_region_type])

                        # Create label text with section type and region number
                        label_text = titles.get(self.current_region_type, self.current_region_type[0].upper())
                        label_text += str(region_index + 1)  # Add region number (1-based)

                        # Check if this region label has already been set in _region_labels_set
                        # If it has, use the existing label to ensure consistency
                        if (self.current_page_index in self._region_labels_set and
                            self.current_region_type in self._region_labels_set[self.current_page_index] and
                            region_index in self._region_labels_set[self.current_page_index][self.current_region_type]):
                            # Use the existing label
                            print(f"[DEBUG] Using existing region label for {self.current_region_type} index {region_index} on page {self.current_page_index + 1}")
                        else:
                            # Mark this label as set
                            if self.current_page_index not in self._region_labels_set:
                                self._region_labels_set[self.current_page_index] = {}
                            if self.current_region_type not in self._region_labels_set[self.current_page_index]:
                                self._region_labels_set[self.current_page_index][self.current_region_type] = {}
                            self._region_labels_set[self.current_page_index][self.current_region_type][region_index] = True
                            print(f"[DEBUG] Marked region label as set: {self.current_region_type} index {region_index} on page {self.current_page_index + 1}")

                        # Add page number to the label for multi-page mode
                        # This ensures each region has a unique identifier across all pages
                        page_num = self.current_page_index + 1  # 1-based page number
                        print(f"[DEBUG] Creating region label for page {page_num}: {label_text}")

                        # Create StandardRegion object - SINGLE FORMAT EVERYWHERE
                        standard_region = self.create_standard_region(
                            rect.x(), rect.y(), rect.width(), rect.height(),
                            self.current_region_type, region_index
                        )

                        self.page_regions[self.current_page_index][self.current_region_type].append(standard_region)

                        # Update the regions dictionary to match page_regions for the current page
                        self.regions = self.page_regions[self.current_page_index]
                        print(f"[DEBUG] Updated regions from page_regions for page {self.current_page_index + 1}")

                        # Extract data for the newly drawn region in multi-page mode
                        # This will be the only extraction call
                        print(f"[DEBUG] Extracting data for newly drawn region {self.current_region_type} index {region_index} on page {self.current_page_index + 1}")

                        # Set flag to indicate we're in a specific region extraction context
                        self._in_specific_region_extraction = True

                        # Set flag to skip extraction update for other pages
                        self._skip_extraction_update = True

                        # Extract data for the specific region
                        self.extract_specific_region(self.current_region_type, region_index)

                        # Keep the flags set until the next mouse press event
                        # This prevents automatic extraction updates that could cause duplication
                    else:
                        # For single-page mode, add to regions
                        region_index = len(self.regions[self.current_region_type])

                        # Create label text with section type and region number
                        label_text = titles.get(self.current_region_type, self.current_region_type[0].upper())
                        label_text += str(region_index + 1)  # Add region number (1-based)

                        # Check if this region label has already been set in _region_labels_set
                        # If it has, use the existing label to ensure consistency
                        if (0 in self._region_labels_set and
                            self.current_region_type in self._region_labels_set[0] and
                            region_index in self._region_labels_set[0][self.current_region_type]):
                            # Use the existing label
                            print(f"[DEBUG] Using existing region label for {self.current_region_type} index {region_index} in single-page mode")
                        else:
                            # Mark this label as set
                            if 0 not in self._region_labels_set:
                                self._region_labels_set[0] = {}
                            if self.current_region_type not in self._region_labels_set[0]:
                                self._region_labels_set[0][self.current_region_type] = {}
                            self._region_labels_set[0][self.current_region_type][region_index] = True
                            print(f"[DEBUG] Marked region label as set: {self.current_region_type} index {region_index} in single-page mode")

                        # For consistency, add page number to the label even in single-page mode
                        # This ensures consistent labeling between single and multi-page modes
                        print(f"[DEBUG] Creating region label for single-page mode: {label_text}")

                        # Create StandardRegion object - SINGLE FORMAT EVERYWHERE
                        standard_region = self.create_standard_region(
                            rect.x(), rect.y(), rect.width(), rect.height(),
                            self.current_region_type, region_index
                        )

                        self.regions[self.current_region_type].append(standard_region)

                    # For single-page mode, we already extracted the data in the multi-page mode branch
                    # So we don't need to extract it again here
                    if not self.multi_page_mode:
                        # Only extract if we're in single-page mode
                        print(f"[DEBUG] Extracting data for newly drawn region {self.current_region_type} index {region_index} in single-page mode")

                        # Set flag to indicate we're in a specific region extraction context
                        self._in_specific_region_extraction = True

                        # Set flag to skip extraction update
                        self._skip_extraction_update = True

                        # Extract data for the specific region
                        self.extract_specific_region(self.current_region_type, region_index)

                        # Keep the flags set until the next mouse press event
                        # This prevents automatic extraction updates that could cause duplication

                    # No need to set _skip_extraction_update here as it's already set above

                    # Don't emit the signal as it would trigger another extraction
                    # self.region_drawn.emit()

            elif self.pdf_label.drawing_mode == 'column' and self.pdf_label.start_pos and pos:
                # Only add column line if we have an active region
                if hasattr(self, 'active_region_type') and hasattr(self, 'active_rect_index'):
                    if self.active_region_type and self.active_rect_index is not None:
                        try:
                            # Get the active rectangle - ENFORCE StandardRegion format
                            region = self.regions[self.active_region_type][self.active_rect_index]

                            # Enforce StandardRegion format - NO backward compatibility
                            from standardized_coordinates import StandardRegion
                            if not isinstance(region, StandardRegion):
                                print(f"[ERROR] Invalid region type: expected StandardRegion, got {type(region)}")
                                return

                            rect = region.rect
                            region_label = region.label
                            print(f"[DEBUG] Drawing column for region with label: {region_label}")

                            # Make sure the x-coordinate stays within the rectangle's bounds
                            x_pos = max(rect.left(), min(pos.x(), rect.right()))

                            # Create vertical line from top to bottom of the rectangle
                            start_point = QPoint(x_pos, rect.top())
                            end_point = QPoint(x_pos, rect.bottom())

                            # Convert active_region_type to RegionType if it's a string
                            region_type = RegionType(self.active_region_type) if isinstance(self.active_region_type, str) else self.active_region_type

                            print(f"[DEBUG] Adding column line at x={x_pos} for region type {region_type.value if hasattr(region_type, 'value') else region_type}")
                            print(f"[DEBUG] Associated with region index {self.active_rect_index}")

                            # Store the column line with the rectangle index
                            column_line = (start_point, end_point, self.active_rect_index)

                            print(f"\n[DEBUG] ===== ADDING COLUMN LINE =====")
                            print(f"[DEBUG] Multi-page mode: {self.multi_page_mode}")
                            print(f"[DEBUG] Current page: {self.current_page_index + 1}")
                            print(f"[DEBUG] Region type: {region_type.value if hasattr(region_type, 'value') else region_type}")
                            print(f"[DEBUG] Region index: {self.active_rect_index}")
                            print(f"[DEBUG] Region label: {region_label}")
                            print(f"[DEBUG] Column position: x={x_pos}")

                            # Print all regions on this page for debugging
                            if self.multi_page_mode and self.current_page_index in self.page_regions:
                                for section_type, regions in self.page_regions[self.current_page_index].items():
                                    region_labels = []
                                    for region in regions:
                                        # Enforce StandardRegion format - NO backward compatibility
                                        if isinstance(region, StandardRegion):
                                            region_labels.append(region.label)
                                        else:
                                            print(f"[ERROR] Invalid region type in debug: expected StandardRegion, got {type(region)}")
                                    print(f"[DEBUG] {section_type} regions on page {self.current_page_index + 1}: {region_labels}")

                            if self.multi_page_mode:
                                # For multi-page mode, only add to page_column_lines
                                if self.current_page_index not in self.page_column_lines:
                                    self.page_column_lines[self.current_page_index] = {}
                                    print(f"[DEBUG] Initialized page_column_lines for page {self.current_page_index + 1}")

                                # Ensure the region type key exists in the dictionary
                                if region_type not in self.page_column_lines[self.current_page_index]:
                                    self.page_column_lines[self.current_page_index][region_type] = []
                                    print(f"[DEBUG] Initialized column lines for region type {region_type.value if hasattr(region_type, 'value') else region_type} in page_column_lines")

                                # Add the column line to page_column_lines
                                self.page_column_lines[self.current_page_index][region_type].append(column_line)
                                print(f"[DEBUG] Added column line to page_column_lines for page {self.current_page_index + 1}")

                                # Update column_lines to match page_column_lines for the current page
                                # This ensures that column_lines always reflects the current page's columns
                                self.column_lines = copy.deepcopy(self.page_column_lines[self.current_page_index])
                                print(f"[DEBUG] Updated column_lines to match page_column_lines for page {self.current_page_index + 1}")
                            else:
                                # For single-page mode, just add to column_lines
                                if region_type not in self.column_lines:
                                    self.column_lines[region_type] = []
                                    print(f"[DEBUG] Initialized column lines for region type {region_type.value if hasattr(region_type, 'value') else region_type} in column_lines")

                                # Add the column line to column_lines
                                self.column_lines[region_type].append(column_line)
                                print(f"[DEBUG] Added column line to column_lines")

                            # Extract data specifically for the region where the column line was drawn
                            # Force extraction even if no changes were detected
                            print(f"[DEBUG] Forcing extraction after column line drawn for region {self.active_rect_index}")

                            # Set _last_extraction_state to None to force extraction
                            self._last_extraction_state = None

                            # Clear cached extraction data for this region to ensure fresh extraction
                            if hasattr(self, '_cached_extraction_data'):
                                section_type = region_type.value if hasattr(region_type, 'value') else region_type
                                if section_type in self._cached_extraction_data:
                                    if isinstance(self._cached_extraction_data[section_type], list):
                                        # If it's a list of DataFrames, clear the specific region
                                        if self.active_rect_index < len(self._cached_extraction_data[section_type]):
                                            self._cached_extraction_data[section_type][self.active_rect_index] = None
                                    else:
                                        # If it's a single DataFrame, clear the entire section
                                        self._cached_extraction_data[section_type] = None
                                    print(f"[DEBUG] Cleared cached extraction data for {section_type} region {self.active_rect_index}")

                            # If we're in multi-page mode, update the stored data for the current page
                            if self.multi_page_mode and hasattr(self, '_all_pages_data') and self._all_pages_data:
                                if self.current_page_index < len(self._all_pages_data) and self._all_pages_data[self.current_page_index] is not None:
                                    # Clear the cached data for this section in the current page
                                    section_type = region_type.value if hasattr(region_type, 'value') else region_type
                                    if section_type in self._all_pages_data[self.current_page_index]:
                                        if isinstance(self._all_pages_data[self.current_page_index][section_type], list):
                                            # If it's a list of DataFrames, clear the specific region
                                            if self.active_rect_index < len(self._all_pages_data[self.current_page_index][section_type]):
                                                self._all_pages_data[self.current_page_index][section_type][self.active_rect_index] = None
                                        else:
                                            # If it's a single DataFrame, clear the entire section
                                            self._all_pages_data[self.current_page_index][section_type] = None
                                        print(f"[DEBUG] Cleared cached extraction data in _all_pages_data for page {self.current_page_index + 1}, {section_type} region {self.active_rect_index}")

                            # Extract the specific region with the new column
                            section_type = region_type.value if hasattr(region_type, 'value') else region_type
                            print(f"[DEBUG] Extracting specific region: {section_type}, index {self.active_rect_index} on page {self.current_page_index + 1}")

                            # Force extraction for the specific region only
                            self._last_extraction_state = None

                            # Set flag to indicate we're in a specific region extraction context
                            self._in_specific_region_extraction = True

                            # Extract data for the specific region
                            self.extract_specific_region(section_type, self.active_rect_index)

                            # Reset the flag
                            self._in_specific_region_extraction = False

                            # Skip the automatic extraction update at the end of handle_mouse_release
                            self._skip_extraction_update = True

                            print(f"[DEBUG] ===== COLUMN LINE ADDED =====\n")
                        except (IndexError, KeyError) as e:
                            print(f"[DEBUG] Error adding column line: {str(e)}")
                    else:
                        print(f"[DEBUG] No active region selected for column line")
                else:
                    print(f"[DEBUG] Missing active_region_type or active_rect_index attributes")

            # Reset cursor based on current mode
            if self.current_region_type:
                self.pdf_label.setCursor(Qt.CrossCursor)
            elif self.drawing_column:
                self.pdf_label.setCursor(Qt.SplitHCursor)
            else:
                self.pdf_label.setCursor(Qt.ArrowCursor)

            # Reset drawing variables
            self.pdf_label.start_pos = None
            self.pdf_label.current_pos = None

            # Update display
            self.pdf_label.update()

            # Update extraction results - force extraction if a column line was just added
            # Skip extraction update if the flag is set
            if hasattr(self, '_skip_extraction_update') and self._skip_extraction_update:
                print(f"[DEBUG] Skipping automatic extraction update due to _skip_extraction_update flag")
                # Reset the flag for next time
                self._skip_extraction_update = False
            elif self.pdf_label.drawing_mode == 'column':
                print(f"[DEBUG] ===== COLUMN LINE ADDED =====")

                # Clear the extraction cache to force re-extraction
                if hasattr(self, '_extraction_cache'):
                    self._extraction_cache = {}
                    print(f"[DEBUG] Cleared extraction cache")

                # Reset extraction state to force re-extraction
                self._last_extraction_state = None

                # Force extraction for all pages in multipage mode
                if self.multi_page_mode:
                    print(f"[DEBUG] Multipage mode detected, forcing extraction for all pages")

                    # Reset all pages data to ensure we re-extract everything
                    if hasattr(self, '_all_pages_data') and self._all_pages_data:
                        self._all_pages_data = [None] * len(self.pdf_document)
                        print(f"[DEBUG] Reset _all_pages_data to force re-extraction")

                    # Debug column lines
                    print(f"[DEBUG] Current page column lines:")
                    if self.current_page_index in self.page_column_lines:
                        for section, lines in self.page_column_lines[self.current_page_index].items():
                            section_name = section.value if hasattr(section, 'value') else section
                            print(f"[DEBUG] {section_name}: {len(lines)} lines")
                    else:
                        print(f"[DEBUG] No column lines for current page {self.current_page_index}")

                    # Extract data for the current page first
                    print(f"[DEBUG] Forcing extraction for current page {self.current_page_index + 1}")
                    self.update_extraction_results(force=True)

                    # Then extract data for all other pages
                    print(f"[DEBUG] Extracting data for all other pages")
                    self.extract_all_pages()
                else:
                    # Single page mode - just force extraction for the current page
                    print(f"[DEBUG] Single page mode, forcing extraction for current page")
                    self.update_extraction_results(force=True)
            else:
                self.update_extraction_results()

    def load_pdf(self, pdf_path=None):
        """Load a PDF file"""
        if not pdf_path:
            # Open file dialog to select PDF
            pdf_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select PDF File",
                "",
                "PDF Files (*.pdf)"
            )

            if not pdf_path:
                return  # User cancelled

        # Load the PDF file
        self.load_pdf_file(pdf_path)

    def _cleanup_pdf_resources(self):
        """Clean up PDF document and related resources to prevent memory leaks"""
        try:
            print("[DEBUG] Cleaning up PDF resources...")

            # Close existing PDF document
            if hasattr(self, 'pdf_document') and self.pdf_document:
                try:
                    self.pdf_document.close()
                    print("[DEBUG] Closed existing PDF document")
                except Exception as e:
                    print(f"[WARNING] Error closing PDF document: {e}")
                finally:
                    self.pdf_document = None

            # Clear PDF label pixmaps to free memory
            if hasattr(self, 'pdf_label') and self.pdf_label:
                try:
                    # Clear both original and scaled pixmaps
                    if hasattr(self.pdf_label, 'original_pixmap'):
                        self.pdf_label.original_pixmap = None
                    if hasattr(self.pdf_label, 'scaled_pixmap'):
                        self.pdf_label.scaled_pixmap = None
                    self.pdf_label.clear()
                    print("[DEBUG] Cleared PDF label pixmaps")
                except Exception as e:
                    print(f"[WARNING] Error clearing PDF label: {e}")

            # Clear extraction cache for previous PDF
            if hasattr(self, 'pdf_path') and self.pdf_path:
                try:
                    from pdf_extraction_utils import clear_extraction_cache_for_pdf
                    clear_extraction_cache_for_pdf(self.pdf_path)
                    print(f"[DEBUG] Cleared extraction cache for: {self.pdf_path}")
                except Exception as e:
                    print(f"[WARNING] Error clearing extraction cache: {e}")

            # Force garbage collection to free memory immediately
            import gc
            collected = gc.collect()
            print(f"[DEBUG] Garbage collection freed {collected} objects")

        except Exception as e:
            print(f"[ERROR] Error in _cleanup_pdf_resources: {e}")

    def _clear_all_caches(self):
        """Clear all caches and cached data to free memory"""
        try:
            print("[DEBUG] Clearing all caches...")

            # Clear extraction cache for current PDF
            if hasattr(self, 'pdf_path') and self.pdf_path:
                try:
                    from pdf_extraction_utils import clear_extraction_cache_for_pdf
                    clear_extraction_cache_for_pdf(self.pdf_path)
                    print(f"[DEBUG] Cleared extraction cache for: {self.pdf_path}")
                except Exception as e:
                    print(f"[WARNING] Error clearing PDF-specific cache: {e}")

            # Clear all extraction caches
            try:
                from pdf_extraction_utils import clear_extraction_cache
                clear_extraction_cache()
                print("[DEBUG] Cleared all extraction caches")
            except Exception as e:
                print(f"[WARNING] Error clearing extraction caches: {e}")

            # Clear instance-specific cached data
            if hasattr(self, '_cached_extraction_data'):
                self._cached_extraction_data = {
                    'header': [],
                    'items': [],
                    'summary': []
                }
                print("[DEBUG] Cleared cached extraction data")

            if hasattr(self, '_all_pages_data'):
                self._all_pages_data = None
                print("[DEBUG] Cleared all pages data")

            if hasattr(self, '_last_extraction_state'):
                self._last_extraction_state = None
                print("[DEBUG] Cleared last extraction state")

            # Clear region data
            if hasattr(self, 'regions'):
                self.regions = {'header': [], 'items': [], 'summary': []}
            if hasattr(self, 'page_regions'):
                self.page_regions = {}
            if hasattr(self, 'page_column_lines'):
                self.page_column_lines = {}
            print("[DEBUG] Cleared region data")

            # Use cache manager if available
            try:
                from cache_manager import get_cache_manager, CACHE_MANAGER_AVAILABLE
                if CACHE_MANAGER_AVAILABLE:
                    cache_manager = get_cache_manager()
                    cache_manager.clear_memory_caches()
                    print("[DEBUG] Used cache manager for memory cleanup")
            except Exception as e:
                print(f"[WARNING] Cache manager not available: {e}")

            print("[DEBUG] All caches cleared successfully")

        except Exception as e:
            print(f"[ERROR] Error in _clear_all_caches: {e}")

    def _shallow_copy_regions(self, regions_dict):
        """Create a shallow copy of regions to reduce memory usage"""
        if not regions_dict:
            return {'header': [], 'items': [], 'summary': []}

        # Use list() constructor for shallow copy instead of deepcopy
        return {
            'header': list(regions_dict.get('header', [])),
            'items': list(regions_dict.get('items', [])),
            'summary': list(regions_dict.get('summary', []))
        }

    def _shallow_copy_column_lines(self, column_lines_dict):
        """Create a shallow copy of column lines to reduce memory usage"""
        if not column_lines_dict:
            return {
                RegionType.HEADER: [], 'header': [],
                RegionType.ITEMS: [], 'items': [],
                RegionType.SUMMARY: [], 'summary': []
            }

        # Use list() constructor for shallow copy instead of deepcopy
        result = {}
        for key, value in column_lines_dict.items():
            if isinstance(value, list):
                result[key] = list(value)  # Shallow copy
            else:
                result[key] = value
        return result

    def _monitor_memory_usage(self):
        """Monitor memory usage and trigger cleanup if necessary"""
        try:
            import psutil
            import os

            # Get current process memory usage
            process = psutil.Process(os.getpid())
            memory_info = process.memory_info()
            memory_mb = memory_info.rss / 1024 / 1024  # Convert to MB

            # Set memory threshold (e.g., 1GB)
            memory_threshold_mb = 1024

            if memory_mb > memory_threshold_mb:
                print(f"[WARNING] High memory usage detected: {memory_mb:.1f}MB > {memory_threshold_mb}MB")
                print("[DEBUG] Triggering automatic memory cleanup...")

                # Clear caches
                self._clear_all_caches()

                # Force garbage collection
                import gc
                collected = gc.collect()
                print(f"[DEBUG] Emergency cleanup freed {collected} objects")

                # Check memory again
                memory_info_after = process.memory_info()
                memory_mb_after = memory_info_after.rss / 1024 / 1024
                print(f"[DEBUG] Memory usage after cleanup: {memory_mb_after:.1f}MB")

            return memory_mb

        except ImportError:
            # psutil not available, skip monitoring
            return 0
        except Exception as e:
            print(f"[WARNING] Error monitoring memory usage: {e}")
            return 0

    def load_pdf_file(self, file_path, skip_dialog=False, source="direct"):
        """
        Load a PDF file and display it

        Args:
            file_path (str): Path to the PDF file
            skip_dialog (bool): Whether to skip the multi-page dialog
            source (str): Source of the PDF loading - "direct" for direct loading,
                         "template_manager" for loading from template manager
        """
        print(f"\n[DEBUG] load_pdf_file called in split_screen_invoice_processor.py")
        print(f"[DEBUG] File path: {file_path}")
        print(f"[DEBUG] Skip dialog: {skip_dialog}")
        print(f"[DEBUG] Loading source: {source}")

        # Clean up previous PDF resources before loading new one
        self._cleanup_pdf_resources()

        # Store the PDF loading source
        self.pdf_loading_source = source

        self.pdf_path = file_path
        self.pdf_document = fitz.open(file_path)
        self.current_page_index = 0

        # Set extraction parameters based on source
        if source == "direct":
            # Use default parameters for direct loading
            self.extraction_params = {
                'header': {'row_tol': 5},
                'items': {'row_tol': 15},
                'summary': {'row_tol': 10},
                'flavor': 'stream',
                'strip_text': '\n'
            }
            print(f"[DEBUG] Using default extraction parameters for direct loading")

        # Initialize page_regions, page_column_lines, and page_configs based on the number of pages
        num_pages = len(self.pdf_document)
        print(f"[DEBUG] Initializing for {num_pages} pages")

        # Reset page_regions - will be populated from template or user drawing
        # Use dictionary format by default for backward compatibility
        self.page_regions = {}

        # Reset page_column_lines - will be populated from template or user drawing
        # Use dictionary format by default for backward compatibility
        self.page_column_lines = {}

        # Reset page_configs - will be populated from template
        self.page_configs = [None] * num_pages

        # Check if this is a multi-page PDF
        if num_pages > 1 and not skip_dialog:
            # Show dialog to ask user if they want to configure all pages
            reply = QMessageBox.question(
                self,
                "Multi-page PDF Detected",
                f"This PDF has {num_pages} pages. Would you like to configure all pages?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                # Enable multi-page mode
                self.multi_page_mode = True
                self.prev_page_btn.show()
                self.next_page_btn.show()
                self.apply_to_remaining_btn.show()
                # Global save template button is always visible

                # Initialize regions and column lines for first page
                self.page_regions[0] = {
                    'header': [],
                    'items': [],
                    'summary': []
                }
                self.page_column_lines[0] = {
                    RegionType.HEADER: [], 'header': [],
                    RegionType.ITEMS: [], 'items': [],
                    RegionType.SUMMARY: [], 'summary': []
                }
            else:
                # User cancelled, close the PDF and return
                self.pdf_document.close()
                self.pdf_document = None
                self.pdf_path = None
                return
        elif num_pages > 1:
            # For multi-page PDFs when applying templates, automatically set multi-page mode
            self.multi_page_mode = True
            self.prev_page_btn.show()
            self.next_page_btn.show()
            self.apply_to_remaining_btn.show()
            # Global save template button is always visible

            # Initialize regions and column lines for first page
            self.page_regions[0] = {
                'header': [],
                'items': [],
                'summary': []
            }
            self.page_column_lines[0] = {
                RegionType.HEADER: [], 'header': [],
                RegionType.ITEMS: [], 'items': [],
                RegionType.SUMMARY: [], 'summary': []
            }
        else:
            # Single-page PDF
            self.multi_page_mode = False
            self.prev_page_btn.hide()
            self.next_page_btn.hide()
            self.apply_to_remaining_btn.hide()
            # Global save template button is always visible

            # Initialize regions and column lines
            self.regions = {'header': [], 'items': [], 'summary': []}
            # Initialize column_lines with both string and enum keys for backward compatibility
            self.column_lines = {
                RegionType.HEADER: [], 'header': [],
                RegionType.ITEMS: [], 'items': [],
                RegionType.SUMMARY: [], 'summary': []
            }

        # Ensure the PDFLabel has direct access to page_configs
        import copy
        self.pdf_label.page_configs = copy.deepcopy(self.page_configs)

        # For multi-page templates, ensure regions and column_lines are initialized from page_configs
        if self.multi_page_mode and isinstance(self.page_configs, list) and len(self.page_configs) > 0:
            # Initialize regions and column_lines for the current page
            self.initialize_from_page_configs()

        # Display the current page
        self.display_current_page()

        # Hide upload area and show PDF label
        self.upload_area.hide()
        self.pdf_label.show()

        # Show PDF controls when PDF is loaded
        self.pdf_controls_container.show()

        # Show scrollbars when PDF is loaded
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        print(f"[DEBUG] Showing scrollbars for PDF display")

        # Reset the JSON tree and clear cached extraction data
        self.json_tree.clear()
        placeholder_item = QTreeWidgetItem(["No data", "No data extracted yet. Draw regions on the PDF to see results."])
        self.json_tree.addTopLevelItem(placeholder_item)

        # Reset cached extraction data
        self._cached_extraction_data = {
            'header': [],
            'items': [],
            'summary': []
        }
        self._last_extraction_state = None

        # Clear extraction cache for this PDF
        from pdf_extraction_utils import clear_extraction_cache_for_pdf
        clear_extraction_cache_for_pdf(self.pdf_path)
        print(f"[DEBUG] Cleared extraction cache for PDF: {self.pdf_path}")

        # Print summary of _all_pages_data after initialization
        self._print_all_pages_data_summary("load_pdf_file")

        # Initialize the invoice2data template with default values
        self.initialize_invoice2data_template()

        # Ensure extraction control buttons remain enabled
        self.retry_btn.setEnabled(True)
        self.adjust_params_btn.setEnabled(True)

    def display_current_page(self):
        """Display the current page of the PDF"""
        if not self.pdf_document:
            return

        # Check if PDF section is visible - if not, just store the state but don't try to render
        pdf_section_width = self.main_splitter.sizes()[0]
        if pdf_section_width < 10:
            self.pdf_section_was_hidden = True
            print(f"[DEBUG] PDF section is hidden. Skipping display update.")
            return

        # Save regions of previous page if in multi-page mode
        if self.multi_page_mode and hasattr(self, 'prev_page_index'):
            if not hasattr(self, 'page_regions'):
                self.page_regions = {}
            if not hasattr(self, 'page_column_lines'):
                self.page_column_lines = {}

            # Use shallow copy to reduce memory usage
            self.page_regions[self.prev_page_index] = self._shallow_copy_regions(self.regions)
            self.page_column_lines[self.prev_page_index] = self._shallow_copy_column_lines(self.column_lines)
            print(f"[DEBUG] Saved regions and column lines for previous page {self.prev_page_index + 1}")

        # Store current page index for next switch
        self.prev_page_index = self.current_page_index

        # Get the current page
        page = self.pdf_document[self.current_page_index]

        # Render the page
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))

        try:
            # Convert PyMuPDF pixmap to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Convert PIL Image to QPixmap
            bytes_io = io.BytesIO()
            img.save(bytes_io, format='PNG')
            qimg = QImage.fromData(bytes_io.getvalue())
            pixmap = QPixmap.fromImage(qimg)

            # Set the pixmap to the PDF label
            self.pdf_label.setPixmap(pixmap)

            # Update coordinate scale factors for standardized coordinate system
            self.update_coordinate_scale_factors()

        finally:
            # CRITICAL: Clean up PyMuPDF resources to prevent memory leaks
            try:
                # Clear the pixmap data to free memory
                pix = None
                # Close the bytes_io buffer
                if 'bytes_io' in locals():
                    bytes_io.close()
                print(f"[DEBUG] Cleaned up PyMuPDF resources for page {self.current_page_index + 1}")
            except Exception as e:
                print(f"[WARNING] Error cleaning up PyMuPDF resources: {e}")

        # Ensure the scroll area shows the entire PDF, including the footer
        # Reset the scroll position to the top
        self.scroll_area.verticalScrollBar().setValue(0)

        # Force the scroll area to update its layout to ensure proper scrolling
        self.scroll_area.updateGeometry()

        # Make sure the PDF label size is properly updated
        self.pdf_label.adjustPixmap()

        # Process events to ensure UI updates immediately
        QApplication.processEvents()

        # Ensure the scroll area can scroll to the end of the document
        self.ensure_full_scroll_range()

        # Show and position the zoom controls
        if hasattr(self, 'zoom_controls'):
            self.zoom_controls.show()
            self.position_zoom_controls()

        # Update regions and column lines based on current page
        if self.multi_page_mode:
            print(f"\n[DEBUG] ===== UPDATING PAGE DISPLAY FOR MULTI-PAGE MODE =====")
            print(f"[DEBUG] Switching to page {self.current_page_index + 1}")

            # CRITICAL FIX: Ensure page_regions are properly initialized from page_configs
            if hasattr(self, 'page_configs') and self.page_configs:
                self.initialize_from_page_configs()
                print(f"[DEBUG] Re-initialized regions from page_configs for page {self.current_page_index + 1}")

            # Get regions and column lines for current page
            if self.current_page_index in self.page_regions:
                # Use shallow copy to reduce memory usage
                self.regions = self._shallow_copy_regions(self.page_regions[self.current_page_index])
                # Initialize any missing region types
                for region_type in ['header', 'items', 'summary']:
                    if region_type not in self.regions:
                        self.regions[region_type] = []
                print(f"[DEBUG] Loaded regions from page_regions for page {self.current_page_index + 1}")
                print(f"[DEBUG] Region counts: header={len(self.regions.get('header', []))}, items={len(self.regions.get('items', []))}, summary={len(self.regions.get('summary', []))}")
            else:
                # Initialize empty regions dictionary for this page
                self.regions = {'header': [], 'items': [], 'summary': []}
                # Also initialize the page_regions entry for this page
                self.page_regions[self.current_page_index] = {'header': [], 'items': [], 'summary': []}
                print(f"[DEBUG] Initialized empty regions for page {self.current_page_index + 1}")

            if self.current_page_index in self.page_column_lines:
                # Use shallow copy to reduce memory usage
                self.column_lines = self._shallow_copy_column_lines(self.page_column_lines[self.current_page_index])
                print(f"[DEBUG] Loaded column lines from page_column_lines for page {self.current_page_index + 1}")

                # Print column line counts for each region type
                for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                    if region_type in self.column_lines:
                        print(f"[DEBUG] Column line count for {region_type.value}: {len(self.column_lines[region_type])}")
                    else:
                        print(f"[DEBUG] No column lines found for {region_type.value}")
            else:
                # Initialize empty column lines dictionary for this page
                self.column_lines = {
                    RegionType.HEADER: [],
                    RegionType.ITEMS: [],
                    RegionType.SUMMARY: []
                }
                # Also initialize the page_column_lines entry for this page
                self.page_column_lines[self.current_page_index] = copy.deepcopy(self.column_lines)
                print(f"[DEBUG] Initialized empty column lines for page {self.current_page_index + 1}")

            # Ensure all region types exist in both dictionaries
            for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                if region_type not in self.column_lines:
                    self.column_lines[region_type] = []
                if self.current_page_index in self.page_column_lines and region_type not in self.page_column_lines[self.current_page_index]:
                    self.page_column_lines[self.current_page_index][region_type] = []

            print(f"[DEBUG] ===== COMPLETED UPDATING PAGE DISPLAY FOR MULTI-PAGE MODE =====\n")

        # Update extraction results using our helper method
        self._update_page_data()

        # Monitor memory usage and cleanup if necessary
        self._monitor_memory_usage()

    def prev_page(self):
        """Go to the previous page"""
        if not self.pdf_document or self.current_page_index <= 0:
            return

        print(f"\n[DEBUG] ===== NAVIGATING TO PREVIOUS PAGE =====")
        print(f"[DEBUG] Current page: {self.current_page_index + 1}")
        print(f"[DEBUG] Multi-page mode: {self.multi_page_mode}")

        # Save current page's regions and column lines
        if self.multi_page_mode:
            print(f"[DEBUG] Saving current page data before navigation")
            if not isinstance(self.page_regions, dict):
                self.page_regions = {}
            if not isinstance(self.page_column_lines, dict):
                self.page_column_lines = {}

            # Save current page state
            self.page_regions[self.current_page_index] = copy.deepcopy(self.regions)
            self.page_column_lines[self.current_page_index] = copy.deepcopy(self.column_lines)
            print(f"[DEBUG] Saved regions and column lines for page {self.current_page_index + 1}")

            # Ensure all region types exist in saved data
            for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                if self.current_page_index in self.page_column_lines and region_type not in self.page_column_lines[self.current_page_index]:
                    self.page_column_lines[self.current_page_index][region_type] = []

        # Go to previous page
        self.current_page_index -= 1
        print(f"[DEBUG] Navigating to page {self.current_page_index + 1}")

        # Display the new page
        self.display_current_page()

        # CRITICAL: Use cached multi-page extraction results if available
        # This ensures that all regions from all pages are preserved
        from pdf_extraction_utils import get_multipage_extraction
        cached_data = get_multipage_extraction(self.pdf_path)
        if cached_data:
            print(f"[DEBUG] Using cached multi-page extraction results for PDF: {self.pdf_path}")
            # Store in instance cache for faster access next time
            self._cached_extraction_data = copy.deepcopy(cached_data)
            # Update the JSON tree with the combined data
            self.update_json_tree(cached_data)
            print(f"[DEBUG] Updated extraction results from cached multi-page data")
        # If no cached data is available, use _all_pages_data if available
        elif hasattr(self, '_all_pages_data') and self._all_pages_data and self.current_page_index < len(self._all_pages_data):
            page_data = self._all_pages_data[self.current_page_index]
            if page_data is not None:
                # Use extract_multi_page_invoice to combine data from all pages
                combined_data = self.extract_multi_page_invoice()
                self.update_json_tree(combined_data)
                print(f"[DEBUG] Updated extraction results from combined data for all pages")
            else:
                # Force extraction for the current page if no stored data is available
                print(f"[DEBUG] No stored data for page {self.current_page_index + 1}, forcing extraction")
                self.update_extraction_results(force=True)
        else:
            # Force extraction for the current page if no stored data is available
            print(f"[DEBUG] No _all_pages_data available for page {self.current_page_index + 1}, forcing extraction")
            self.update_extraction_results(force=True)

        print(f"[DEBUG] ===== COMPLETED NAVIGATION TO PREVIOUS PAGE =====\n")

    def next_page(self):
        """Go to the next page"""
        if not self.pdf_document or self.current_page_index >= len(self.pdf_document) - 1:
            return

        print(f"\n[DEBUG] ===== NAVIGATING TO NEXT PAGE =====")
        print(f"[DEBUG] Current page: {self.current_page_index + 1}")
        print(f"[DEBUG] Multi-page mode: {self.multi_page_mode}")

        # Save current page's regions and column lines
        if self.multi_page_mode:
            print(f"[DEBUG] Saving current page data before navigation")
            if not isinstance(self.page_regions, dict):
                self.page_regions = {}
            if not isinstance(self.page_column_lines, dict):
                self.page_column_lines = {}

            # Save current page state
            self.page_regions[self.current_page_index] = copy.deepcopy(self.regions)
            self.page_column_lines[self.current_page_index] = copy.deepcopy(self.column_lines)
            print(f"[DEBUG] Saved regions and column lines for page {self.current_page_index + 1}")

            # Ensure all region types exist in saved data
            for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                if self.current_page_index in self.page_column_lines and region_type not in self.page_column_lines[self.current_page_index]:
                    self.page_column_lines[self.current_page_index][region_type] = []

        # Go to next page
        self.current_page_index += 1
        print(f"[DEBUG] Navigating to page {self.current_page_index + 1}")

        # Display the new page
        self.display_current_page()

        # CRITICAL: Use cached multi-page extraction results if available
        # This ensures that all regions from all pages are preserved
        from pdf_extraction_utils import get_multipage_extraction
        cached_data = get_multipage_extraction(self.pdf_path)
        if cached_data:
            print(f"[DEBUG] Using cached multi-page extraction results for PDF: {self.pdf_path}")
            # Store in instance cache for faster access next time
            self._cached_extraction_data = copy.deepcopy(cached_data)
            # Update the JSON tree with the combined data
            self.update_json_tree(cached_data)
            print(f"[DEBUG] Updated extraction results from cached multi-page data")
        # If no cached data is available, use _all_pages_data if available
        elif hasattr(self, '_all_pages_data') and self._all_pages_data and self.current_page_index < len(self._all_pages_data):
            page_data = self._all_pages_data[self.current_page_index]
            if page_data is not None:
                # Use extract_multi_page_invoice to combine data from all pages
                combined_data = self.extract_multi_page_invoice()
                self.update_json_tree(combined_data)
                print(f"[DEBUG] Updated extraction results from combined data for all pages")
            else:
                # Force extraction for the current page if no stored data is available
                print(f"[DEBUG] No stored data for page {self.current_page_index + 1}, forcing extraction")
                self.update_extraction_results(force=True)
        else:
            # Force extraction for the current page if no stored data is available
            print(f"[DEBUG] No _all_pages_data available for page {self.current_page_index + 1}, forcing extraction")
            self.update_extraction_results(force=True)

        print(f"[DEBUG] ===== COMPLETED NAVIGATION TO NEXT PAGE =====\n")

    def apply_to_remaining_pages(self):
        """Apply current page's regions and column lines to all remaining pages"""
        if not self.pdf_document or not self.multi_page_mode:
            return

        print(f"\n[DEBUG] ===== APPLYING REGIONS AND COLUMNS TO REMAINING PAGES =====")
        print(f"[DEBUG] Current page: {self.current_page_index + 1}")
        print(f"[DEBUG] Total pages: {len(self.pdf_document)}")

        # Save current page's regions and column lines
        print(f"[DEBUG] Saving current page data before applying to other pages")

        # Handle both list and dictionary formats for page_regions
        if isinstance(self.page_regions, dict):
            # Create a deep copy to avoid reference issues
            self.page_regions[self.current_page_index] = copy.deepcopy(self.regions)
            print(f"[DEBUG] Saved regions to page_regions dictionary for page {self.current_page_index + 1}")
        elif isinstance(self.page_regions, list):
            # Ensure the list is long enough
            while len(self.page_regions) <= self.current_page_index:
                self.page_regions.append({})
            # Create a deep copy to avoid reference issues
            self.page_regions[self.current_page_index] = copy.deepcopy(self.regions)
            print(f"[DEBUG] Saved regions to page_regions list for page {self.current_page_index + 1}")
        else:
            print(f"[DEBUG] Unexpected page_regions type: {type(self.page_regions)}")

        # Handle both list and dictionary formats for page_column_lines
        if isinstance(self.page_column_lines, dict):
            # Make a deep copy of column_lines to page_column_lines
            self.page_column_lines[self.current_page_index] = copy.deepcopy(self.column_lines)
            print(f"[DEBUG] Saved column lines to page_column_lines dictionary for page {self.current_page_index + 1}")

            # Print column line counts for each region type
            for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                if region_type in self.column_lines:
                    print(f"[DEBUG] Column line count for {region_type.value}: {len(self.column_lines[region_type])}")

                    # Verify the column lines were properly saved
                    if region_type in self.page_column_lines[self.current_page_index]:
                        saved_count = len(self.page_column_lines[self.current_page_index][region_type])
                        print(f"[DEBUG] Verified {saved_count} column lines saved for {region_type.value}")

                        # If counts don't match, something went wrong
                        if saved_count != len(self.column_lines[region_type]):
                            print(f"[WARNING] Column line count mismatch for {region_type.value}: {saved_count} saved vs {len(self.column_lines[region_type])} in memory")
        elif isinstance(self.page_column_lines, list):
            # Ensure the list is long enough
            while len(self.page_column_lines) <= self.current_page_index:
                self.page_column_lines.append({})
            self.page_column_lines[self.current_page_index] = copy.deepcopy(self.column_lines)
            print(f"[DEBUG] Saved column lines to page_column_lines list for page {self.current_page_index + 1}")
        else:
            print(f"[DEBUG] Unexpected page_column_lines type: {type(self.page_column_lines)}")

        # Ensure all region types exist in the saved data
        for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
            if self.current_page_index in self.page_column_lines and region_type not in self.page_column_lines[self.current_page_index]:
                self.page_column_lines[self.current_page_index][region_type] = []

        # Apply to all remaining pages
        print(f"[DEBUG] Applying to remaining pages ({len(self.pdf_document) - self.current_page_index - 1} pages)")
        for i in range(self.current_page_index + 1, len(self.pdf_document)):
            print(f"[DEBUG] Applying to page {i + 1}")

            # Handle both list and dictionary formats for page_regions
            if isinstance(self.page_regions, dict):
                # Create a deep copy to avoid reference issues
                self.page_regions[i] = copy.deepcopy(self.regions)
                print(f"[DEBUG] Applied regions to page {i + 1}")
            elif isinstance(self.page_regions, list):
                # Ensure the list is long enough
                while len(self.page_regions) <= i:
                    self.page_regions.append({})
                # Create a deep copy to avoid reference issues
                self.page_regions[i] = copy.deepcopy(self.regions)
                print(f"[DEBUG] Applied regions to page {i + 1}")

            # Handle both list and dictionary formats for page_column_lines
            if isinstance(self.page_column_lines, dict):
                # Make a deep copy of column_lines to page_column_lines for this page
                self.page_column_lines[i] = copy.deepcopy(self.column_lines)
                print(f"[DEBUG] Applied column lines to page {i + 1}")

                # Print column line counts for each region type
                for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                    if region_type in self.column_lines:
                        # Verify the column lines were properly copied
                        if region_type in self.page_column_lines[i]:
                            copied_count = len(self.page_column_lines[i][region_type])
                            print(f"[DEBUG] Applied {copied_count} column lines for {region_type.value} to page {i + 1}")

                            # If counts don't match, something went wrong
                            if copied_count != len(self.column_lines[region_type]):
                                print(f"[WARNING] Column line count mismatch for {region_type.value} on page {i + 1}: {copied_count} copied vs {len(self.column_lines[region_type])} in source")
            elif isinstance(self.page_column_lines, list):
                # Ensure the list is long enough
                while len(self.page_column_lines) <= i:
                    self.page_column_lines.append({})
                self.page_column_lines[i] = copy.deepcopy(self.column_lines)
                print(f"[DEBUG] Applied column lines to page {i + 1}")

            # Ensure all region types exist in the copied data
            for region_type in [RegionType.HEADER, RegionType.ITEMS, RegionType.SUMMARY]:
                if i in self.page_column_lines and region_type not in self.page_column_lines[i]:
                    self.page_column_lines[i][region_type] = []

        # Extract data for all pages after applying the template
        try:
            print(f"\n[DEBUG] Extracting data for all pages after applying template")
            # Extract data for each page
            all_pages_data = []
            for page_idx in range(len(self.pdf_document)):
                print(f"\n[DEBUG] Extracting data for page {page_idx + 1}")
                header_df, items_df, summary_df = self.extract_page_data(page_idx)
                all_pages_data.append({
                    'header': header_df,
                    'items': items_df,
                    'summary': summary_df
                })
                print(f"[DEBUG] Extracted data for page {page_idx + 1}")

            # Store the extracted data
            self._all_pages_data = all_pages_data

            # Update the current page's extraction results
            self.update_extraction_results(force=True)

            # Show confirmation message
            QMessageBox.information(
                self,
                "Applied to Remaining Pages",
                f"The current configuration has been applied to all remaining pages and data has been extracted."
            )
        except Exception as e:
            print(f"Error extracting data after applying template: {str(e)}")
            import traceback
            traceback.print_exc()

            # Show error message but still confirm the template was applied
            QMessageBox.information(
                self,
                "Applied to Remaining Pages",
                f"The current configuration has been applied to all remaining pages, but there was an error extracting data: {str(e)}"
            )

    def combine_multipage_data(self):
        """Combine data from all pages in multipage mode using the consolidated method

        In multi-page PDF extraction mode, we preserve region data from all pages without duplicate checking,
        as the same region can appear on different pages. This method ensures that all page data is shown
        in extraction results, not just current page data.

        Returns:
            dict: Dictionary containing combined data from all pages
        """
        if not self.multi_page_mode or not self.pdf_document:
            return

        print(f"[DEBUG] Combining data from all pages in multipage mode using extract_multi_page_invoice")
        print(f"[DEBUG] Ensuring ALL data from ALL pages is preserved without duplicate checking")

    def navigate_to_page(self, page_index):
        """Navigate to a specific page in the PDF"""
        try:
            print(f"\n[DEBUG] ===== NAVIGATING TO PAGE {page_index+1} =====")
            print(f"[DEBUG] Current page: {self.current_page_index+1}")
            print(f"[DEBUG] Multi-page mode: {hasattr(self, 'multi_page_mode') and self.multi_page_mode}")

            # Save current page data before navigation
            print(f"[DEBUG] Saving current page data before navigation")
            self._save_current_page_data()

            # Update current page index
            self.current_page_index = page_index

            # Display the new page
            self.display_current_page()

            # Update page data while preserving all pages' data
            self._update_page_data(page_index)

            # Update UI elements
            if hasattr(self, 'page_label'):
                self.page_label.setText(f"Page {page_index + 1} of {self.pdf_document.pageCount()}")

            print(f"[DEBUG] ===== COMPLETED NAVIGATION TO PAGE {page_index+1} =====")

        except Exception as e:
            print(f"[ERROR] Error navigating to page {page_index+1}: {str(e)}")
            import traceback
            traceback.print_exc()

        # First, make sure we have data for all pages
        for page_idx in range(len(self.pdf_document)):
            # Check if we need to extract data for this page
            if not hasattr(self, '_all_pages_data') or not self._all_pages_data or len(self._all_pages_data) <= page_idx or self._all_pages_data[page_idx] is None:
                print(f"[DEBUG] Extracting data for page {page_idx + 1} before combining")
                # Extract data for this page
                page_data = self.multipage_extract_page_data(page_idx)

                # Initialize _all_pages_data if needed
                if not hasattr(self, '_all_pages_data') or not self._all_pages_data:
                    self._all_pages_data = [None] * len(self.pdf_document)
                elif len(self._all_pages_data) < len(self.pdf_document):
                    # Extend the list if needed
                    self._all_pages_data.extend([None] * (len(self.pdf_document) - len(self._all_pages_data)))

                # Store the extracted data
                self._all_pages_data[page_idx] = page_data
                print(f"[DEBUG] Stored extraction data for page {page_idx + 1}")

        # Use the consolidated method to extract and combine data from all pages
        combined_data = self.extract_multi_page_invoice()

        # The metadata and page_info are already included in the combined_data from extract_multi_page_invoice

        # Print summary of combined data to verify all regions are preserved
        print(f"[DEBUG] Combined data summary:")

        # Check header data
        if isinstance(combined_data['header'], list):
            print(f"[DEBUG] header: List with {len(combined_data['header'])} items")
            for i, df in enumerate(combined_data['header']):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    if 'region_label' in df.columns:
                        region_labels = df['region_label'].tolist()
                        print(f"[DEBUG] header[{i}] region labels: {region_labels}")
                    if 'page_number' in df.columns:
                        page_numbers = df['page_number'].unique().tolist()
                        print(f"[DEBUG] header[{i}] page numbers: {page_numbers}")
                    print(f"[DEBUG] header[{i}] data shape: {df.shape}")
                    print(f"[DEBUG] header[{i}] row count: {len(df)}")
        elif isinstance(combined_data['header'], pd.DataFrame) and not combined_data['header'].empty:
            print(f"[DEBUG] header: DataFrame with {len(combined_data['header'])} rows")
            if 'region_label' in combined_data['header'].columns:
                region_labels = combined_data['header']['region_label'].tolist()
                print(f"[DEBUG] header region labels: {region_labels}")
            if 'page_number' in combined_data['header'].columns:
                page_numbers = combined_data['header']['page_number'].unique().tolist()
                print(f"[DEBUG] header page numbers: {page_numbers}")
        else:
            print(f"[DEBUG] header: None or empty")

        # Check items data
        if isinstance(combined_data['items'], list):
            print(f"[DEBUG] items: List with {len(combined_data['items'])} items")
            for i, df in enumerate(combined_data['items']):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    if 'region_label' in df.columns:
                        region_labels = df['region_label'].tolist()
                        print(f"[DEBUG] items[{i}] region labels: {region_labels}")
                    if 'page_number' in df.columns:
                        page_numbers = df['page_number'].unique().tolist()
                        print(f"[DEBUG] items[{i}] page numbers: {page_numbers}")
                    print(f"[DEBUG] items[{i}] data shape: {df.shape}")
                    print(f"[DEBUG] items[{i}] row count: {len(df)}")
        elif isinstance(combined_data['items'], pd.DataFrame) and not combined_data['items'].empty:
            print(f"[DEBUG] items: DataFrame with {len(combined_data['items'])} rows")
            if 'region_label' in combined_data['items'].columns:
                region_labels = combined_data['items']['region_label'].tolist()
                print(f"[DEBUG] items region labels: {region_labels}")
            if 'page_number' in combined_data['items'].columns:
                page_numbers = combined_data['items']['page_number'].unique().tolist()
                print(f"[DEBUG] items page numbers: {page_numbers}")
        else:
            print(f"[DEBUG] items: None or empty")

        # Check summary data
        if isinstance(combined_data['summary'], list):
            print(f"[DEBUG] summary: List with {len(combined_data['summary'])} items")
            for i, df in enumerate(combined_data['summary']):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    if 'region_label' in df.columns:
                        region_labels = df['region_label'].tolist()
                        print(f"[DEBUG] summary[{i}] region labels: {region_labels}")
                    if 'page_number' in df.columns:
                        page_numbers = df['page_number'].unique().tolist()
                        print(f"[DEBUG] summary[{i}] page numbers: {page_numbers}")
                    print(f"[DEBUG] summary[{i}] data shape: {df.shape}")
                    print(f"[DEBUG] summary[{i}] row count: {len(df)}")
        elif isinstance(combined_data['summary'], pd.DataFrame) and not combined_data['summary'].empty:
            print(f"[DEBUG] summary: DataFrame with {len(combined_data['summary'])} rows")
            if 'region_label' in combined_data['summary'].columns:
                region_labels = combined_data['summary']['region_label'].tolist()
                print(f"[DEBUG] summary region labels: {region_labels}")
            if 'page_number' in combined_data['summary'].columns:
                page_numbers = combined_data['summary']['page_number'].unique().tolist()
                print(f"[DEBUG] summary page numbers: {page_numbers}")
        else:
            print(f"[DEBUG] summary: None or empty")

        # Store the combined data
        self._cached_extraction_data = copy.deepcopy(combined_data)
        print(f"[DEBUG] Stored combined data in _cached_extraction_data: {list(combined_data.keys())}")

        # Update the JSON tree with the combined data
        print(f"[DEBUG] Using combined data from all pages for display")
        self.update_json_tree(combined_data)
        print(f"[DEBUG] Completed combining data from all pages with ALL data from ALL pages preserved")

        return combined_data

    def update_region_labels_with_page_numbers(self, df):
        """Update region labels in a DataFrame to include page numbers

        In multi-page PDF extraction mode, we preserve region data from all pages without duplicate checking,
        as the same region can appear on different pages. This method ensures that region labels include
        page numbers to differentiate between regions with the same name on different pages.

        Args:
            df (pd.DataFrame): DataFrame containing region data

        Returns:
            pd.DataFrame: DataFrame with updated region labels
        """
        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            return df

        if 'region_label' not in df.columns:
            return df

        # Make a copy to avoid modifying the original
        df_copy = df.copy()

        # Create a mapping of old labels to new labels with page numbers
        label_mapping = {}
        for idx, row in df_copy.iterrows():
            old_label = row['region_label']

            # Get page number from the appropriate column
            if '_page_number' in df_copy.columns:
                page_num = row['_page_number']
            elif 'page_number' in df_copy.columns:
                page_num = row['page_number']
            else:
                page_num = 1  # Default to page 1 if no page number column exists

            # Check if the label already has a page number suffix
            if '_P' in old_label:
                # Keep the existing label
                new_label = old_label
            else:
                # Add page number suffix
                new_label = f"{old_label}_P{page_num}"

            label_mapping[idx] = new_label

        # Update the region labels
        for idx, new_label in label_mapping.items():
            df_copy.at[idx, 'region_label'] = new_label

        # Print debug information to verify all regions are preserved
        if 'region_label' in df_copy.columns:
            region_labels = df_copy['region_label'].tolist()
            print(f"[DEBUG] Updated region labels: {region_labels}")

        if '_page_number' in df_copy.columns:
            page_numbers = df_copy['_page_number'].unique().tolist()
            print(f"[DEBUG] Page numbers in data: {page_numbers}")
        elif 'page_number' in df_copy.columns:
            page_numbers = df_copy['page_number'].unique().tolist()
            print(f"[DEBUG] Page numbers in data: {page_numbers}")

        # Count unique base region names (without page numbers)
        base_region_names = set()
        for label in df_copy['region_label']:
            if isinstance(label, str):
                parts = label.split('_')
                if len(parts) > 0:
                    base_region_names.add(parts[0])

        print(f"[DEBUG] Found {len(base_region_names)} unique base region names: {base_region_names}")
        print(f"[DEBUG] Updated region labels in combined DataFrame while preserving ALL regions from all pages")

        return df_copy

    def extract_all_pages(self):
        """Extract data for all pages in multipage mode using the consolidated method

        In multi-page PDF extraction mode, we preserve region data from all pages without duplicate checking,
        as the same region can appear on different pages. This method ensures that all page data is shown
        in extraction results, not just current page data.
        """
        if not self.multi_page_mode or not self.pdf_document:
            return

        print(f"[DEBUG] Extracting data for all pages in multipage mode using extract_multi_page_invoice")
        print(f"[DEBUG] Ensuring ALL data from ALL pages is preserved without duplicate checking")

        # First, make sure we have data for all pages
        for page_idx in range(len(self.pdf_document)):
            # Check if we need to extract data for this page
            if not hasattr(self, '_all_pages_data') or not self._all_pages_data or len(self._all_pages_data) <= page_idx or self._all_pages_data[page_idx] is None:
                print(f"[DEBUG] Extracting data for page {page_idx + 1}")
                # Extract data for this page
                page_data = self.multipage_extract_page_data(page_idx)

                # Initialize _all_pages_data if needed
                if not hasattr(self, '_all_pages_data') or not self._all_pages_data:
                    self._all_pages_data = [None] * len(self.pdf_document)
                elif len(self._all_pages_data) < len(self.pdf_document):
                    # Extend the list if needed
                    self._all_pages_data.extend([None] * (len(self.pdf_document) - len(self._all_pages_data)))

                # Store the extracted data
                self._all_pages_data[page_idx] = page_data
                print(f"[DEBUG] Stored extraction data for page {page_idx + 1}")

        # Use the consolidated method to extract and combine data from all pages
        combined_data = self.extract_multi_page_invoice()

        # Print summary of combined data to verify all regions are preserved
        for section in ['header', 'items', 'summary']:
            if section in combined_data and isinstance(combined_data[section], list):
                print(f"[DEBUG] Combined {section} is a list with {len(combined_data[section])} items")
                for i, df in enumerate(combined_data[section]):
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        if 'region_label' in df.columns:
                            region_labels = df['region_label'].tolist()
                            print(f"[DEBUG] Combined {section}[{i}] region labels: {region_labels}")
                        if 'page_number' in df.columns:
                            page_numbers = df['page_number'].unique().tolist()
                            print(f"[DEBUG] Combined {section}[{i}] page numbers: {page_numbers}")
                        print(f"[DEBUG] Combined {section}[{i}] data shape: {df.shape}")
                        print(f"[DEBUG] Combined {section}[{i}] row count: {len(df)}")

        # Update the JSON tree with the combined data
        self.update_json_tree(combined_data)

        # Store the combined data
        if combined_data and any(section for section in combined_data.values() if section is not None and (not isinstance(section, list) or len(section) > 0)):
            self._cached_extraction_data = copy.deepcopy(combined_data)
            print(f"[DEBUG] Stored combined data in _cached_extraction_data: {list(combined_data.keys())}")

        print(f"[DEBUG] Completed extraction for all pages with ALL data from ALL pages preserved")

    def update_extraction_results(self, force=False):
        """Update the extraction results based on current regions and column lines

        Args:
            force (bool): If True, force extraction even if it would normally be skipped
        """
        if not self.pdf_path:
            return

        # Check if we should skip extraction due to the _skip_extraction_update flag
        if hasattr(self, '_skip_extraction_update') and self._skip_extraction_update:
            print(f"[DEBUG] Skipping extraction update due to _skip_extraction_update flag")
            # Reset the flag for next time
            self._skip_extraction_update = False

            # Make sure we update the display with the cached data
            if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data:
                print(f"[DEBUG] Updating display with cached extraction data")
                self.update_json_tree(self._cached_extraction_data)
            return

        # If we have regions, extract data
        if any(len(rects) > 0 for rects in self.regions.values()):
            try:
                # Check if we should skip extraction (to avoid duplicate calls)
                if not force and hasattr(self, '_last_extraction_state') and self._last_extraction_state is not None:
                    # Compare current state with last extraction state
                    current_state = self._get_extraction_state()
                    if current_state == self._last_extraction_state:
                        print(f"[DEBUG] Skipping duplicate extraction (no changes detected)")

                        # Make sure we update the display with the cached data
                        if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data:
                            print(f"[DEBUG] Updating display with cached extraction data")
                            self.update_json_tree(self._cached_extraction_data)
                        return

                # If force is True or we're in multipage mode, always print a message
                if force or self.multi_page_mode:
                    print(f"[DEBUG] Forcing extraction: force={force}, multi_page_mode={self.multi_page_mode}")

                # Initialize _table_areas if it doesn't exist
                if not hasattr(self, '_table_areas'):
                    self._table_areas = {
                        'header': [],
                        'items': [],
                        'summary': []
                    }

                # Extract data for the current page
                print(f"[DEBUG] Performing full extraction for page {self.current_page_index + 1}")
                header_df, items_df, summary_df = self.extract_page_data(self.current_page_index)

                # Save the current extraction state
                self._last_extraction_state = self._get_extraction_state()

                # Create page data dictionary
                page_data = {
                    'header': header_df,
                    'items': items_df,
                    'summary': summary_df
                }

                # Store the extracted data in _all_pages_data
                if not hasattr(self, '_all_pages_data') or not self._all_pages_data:
                    self._all_pages_data = [None] * len(self.pdf_document)

                # Update the data for the current page
                if self.current_page_index < len(self._all_pages_data):
                    self._all_pages_data[self.current_page_index] = page_data
                    print(f"[DEBUG] Stored extraction data for page {self.current_page_index + 1}")

                # Update the JSON tree with the extracted data
                if self.multi_page_mode and hasattr(self, 'pdf_document') and len(self.pdf_document) > 1:
                    # For multi-page mode, use extract_multi_page_invoice to combine data from all pages
                    print(f"[DEBUG] In multi-page mode, using extract_multi_page_invoice to combine data from all pages")

                    # Make sure we have data for all pages before combining
                    for page_idx in range(len(self.pdf_document)):
                        if page_idx >= len(self._all_pages_data) or self._all_pages_data[page_idx] is None:
                            print(f"[DEBUG] Extracting data for page {page_idx + 1} before combining")
                            # Extract data for this page
                            header_df, items_df, summary_df = self.extract_page_data(page_idx)

                            # Store the extracted data in _all_pages_data
                            page_data = {
                                'header': header_df,
                                'items': items_df,
                                'summary': summary_df
                            }
                            self._all_pages_data[page_idx] = page_data
                            print(f"[DEBUG] Stored extraction data for page {page_idx + 1}")

                    # CRITICAL: Use cached multi-page extraction results if available
                    # This ensures that all regions from all pages are preserved
                    from pdf_extraction_utils import get_multipage_extraction
                    cached_data = get_multipage_extraction(self.pdf_path)
                    if cached_data:
                        print(f"[DEBUG] Using cached multi-page extraction results for PDF: {self.pdf_path}")
                        # Make a deep copy to avoid modifying the original
                        combined_data = copy.deepcopy(cached_data)
                    else:
                        # If no cached data is available, extract and combine data from all pages
                        combined_data = self.extract_multi_page_invoice()
                    print(f"[DEBUG] Combined data from all pages")

                    # Add metadata
                    if self.pdf_path:
                        # If metadata already exists, preserve the original creation_date
                        if 'metadata' in combined_data and 'creation_date' in combined_data['metadata']:
                            creation_date = combined_data['metadata']['creation_date']
                            print(f"[DEBUG] Metadata already exists in combined data in update_extraction_results, preserving original creation_date: {creation_date}")
                        else:
                            creation_date = datetime.datetime.now().isoformat()
                            print(f"[DEBUG] Added metadata to combined data in update_extraction_results with creation_date: {creation_date}")

                        combined_data['metadata'] = {
                            'filename': os.path.basename(self.pdf_path),
                            'page_count': len(self.pdf_document),
                            'template_type': 'multi',  # Always 'multi' for multi-page PDFs
                            'creation_date': creation_date
                        }

                    # Store the combined data in _cached_extraction_data for future use
                    if combined_data:
                        self._cached_extraction_data = combined_data.copy()
                        print(f"[DEBUG] Stored combined data in _cached_extraction_data: {list(combined_data.keys())}")

                        # Print summary of combined data to verify all regions are preserved
                        for section in ['header', 'items', 'summary']:
                            if section in combined_data and isinstance(combined_data[section], pd.DataFrame) and not combined_data[section].empty:
                                if 'region_label' in combined_data[section].columns:
                                    region_labels = combined_data[section]['region_label'].tolist()
                                    print(f"[DEBUG] Combined {section} region labels: {region_labels}")
                                if 'page_number' in combined_data[section].columns:
                                    page_numbers = combined_data[section]['page_number'].unique().tolist()
                                    print(f"[DEBUG] Combined {section} page numbers: {page_numbers}")
                                print(f"[DEBUG] Combined {section} data shape: {combined_data[section].shape}")
                                print(f"[DEBUG] Combined {section} row count: {len(combined_data[section])}")
                            elif section in combined_data and isinstance(combined_data[section], list) and combined_data[section]:
                                print(f"[DEBUG] Combined {section} is a list with {len(combined_data[section])} items")
                                for i, df in enumerate(combined_data[section]):
                                    if isinstance(df, pd.DataFrame) and not df.empty:
                                        if 'region_label' in df.columns:
                                            region_labels = df['region_label'].tolist()
                                            print(f"[DEBUG] Combined {section}[{i}] region labels: {region_labels}")
                                        if 'page_number' in df.columns:
                                            page_numbers = df['page_number'].unique().tolist()
                                            print(f"[DEBUG] Combined {section}[{i}] page numbers: {page_numbers}")
                                        print(f"[DEBUG] Combined {section}[{i}] data shape: {df.shape}")
                                        print(f"[DEBUG] Combined {section}[{i}] row count: {len(df)}")
                            else:
                                print(f"[DEBUG] Combined {section} is empty or None")

                    # Update the display
                    self.update_json_tree(combined_data)
                else:
                    # For single-page mode, just use the current page data
                    # Store the page data in _cached_extraction_data for future use
                    if page_data:
                        self._cached_extraction_data = page_data.copy()
                        print(f"[DEBUG] Stored page data in _cached_extraction_data: {list(page_data.keys())}")

                    # Update the display
                    self.update_json_tree(page_data)
            except Exception as e:
                # Handle extraction errors gracefully
                print(f"Error updating extraction results: {str(e)}")
                import traceback
                traceback.print_exc()

                # Show error message in the JSON tree
                self.json_tree.clear()
                error_item = QTreeWidgetItem(["Error", f"Error extracting data: {str(e)}\n\nTry adjusting the regions or column lines."])
                self.json_tree.addTopLevelItem(error_item)

                # Make sure we update the display with the cached data if available
                if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data:
                    print(f"[DEBUG] Updating display with cached extraction data after error")
                    self.update_json_tree(self._cached_extraction_data)



    def _get_current_json_data(self):
        """Get the current JSON data from the tree view

        For multi-page PDFs, this will combine data from all pages by section.
        In multi-page PDF extraction mode, we preserve region data from all pages without duplicate checking,
        as the same region can appear on different pages.

        Returns:
            dict: The current JSON data with combined data from all pages if in multi-page mode
        """
        try:
            # Check if we're in multi-page mode and need to combine data from all pages
            if self.multi_page_mode and hasattr(self, 'pdf_document') and len(self.pdf_document) > 1:
                print(f"[DEBUG] Multi-page mode detected in _get_current_json_data, combining data from all pages")
                print(f"[DEBUG] Preserving ALL regions from all pages without duplicate checking")

                # Extract combined data from all pages
                combined_data = self.extract_multi_page_invoice()

                # Add metadata
                if self.pdf_path:
                    combined_data['metadata'] = {
                        'filename': os.path.basename(self.pdf_path),
                        'page_count': len(self.pdf_document),
                        'template_type': 'multi',
                        'creation_date': datetime.datetime.now().isoformat()
                    }

                # Print summary of combined data to verify all regions are preserved
                for section in ['header', 'items', 'summary']:
                    if section in combined_data and isinstance(combined_data[section], pd.DataFrame) and not combined_data[section].empty:
                        if 'region_label' in combined_data[section].columns:
                            region_labels = combined_data[section]['region_label'].tolist()
                            print(f"[DEBUG] Combined {section} region labels: {region_labels}")
                        if 'page_number' in combined_data[section].columns:
                            page_numbers = combined_data[section]['page_number'].unique().tolist()
                            print(f"[DEBUG] Combined {section} page numbers: {page_numbers}")
                        print(f"[DEBUG] Combined {section} data shape: {combined_data[section].shape}")
                        print(f"[DEBUG] Combined {section} row count: {len(combined_data[section])}")

                return combined_data
            else:
                # Single page mode - use cached data
                if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data:
                    print(f"[DEBUG] Using cached extraction data")

                    # Create a copy of the cached data
                    data = copy.deepcopy(self._cached_extraction_data)

                    # Add metadata
                    if self.pdf_path:
                        data['metadata'] = {
                            'filename': os.path.basename(self.pdf_path),
                            'page_count': len(self.pdf_document) if hasattr(self, 'pdf_document') and self.pdf_document else 1,
                            'template_type': 'single',
                            'creation_date': datetime.datetime.now().isoformat()
                        }

                    return data
                else:
                    # Initialize with empty data
                    data = {
                        'header': None,
                        'items': [],
                        'summary': None
                    }

                    # Add metadata
                    if hasattr(self, 'pdf_path') and self.pdf_path:
                        data['metadata'] = {
                            'filename': os.path.basename(self.pdf_path),
                            'page_count': len(self.pdf_document) if hasattr(self, 'pdf_document') and self.pdf_document else 1,
                            'template_type': 'single',
                            'creation_date': datetime.datetime.now().isoformat()
                        }

                    print(f"[DEBUG] No cached data available, returning empty structure")
                    return data
        except Exception as e:
            print(f"[ERROR] Failed to get current JSON data: {str(e)}")
            return {
                'metadata': {
                    'filename': os.path.basename(self.pdf_path) if hasattr(self, 'pdf_path') and self.pdf_path else 'Unknown',
                    'page_count': len(self.pdf_document) if hasattr(self, 'pdf_document') and self.pdf_document else 1,
                    'template_type': 'multi' if hasattr(self, 'multi_page_mode') and self.multi_page_mode else 'single',
                    'creation_date': datetime.datetime.now().isoformat()
                },
                'header': None,
                'items': [],
                'summary': None
            }

    def _get_extraction_state(self):
        """Get a hashable representation of the current extraction state

        Returns:
            tuple: A tuple containing the current regions and column lines
        """
        try:
            # Convert regions to a hashable format
            regions_hash = tuple()
            for section, rects in self.regions.items():
                section_rects = []
                for rect in rects:
                    try:
                        # Check if the rect is a dictionary with a 'rect' key
                        if isinstance(rect, dict) and 'rect' in rect:
                            rect_obj = rect['rect']
                            if hasattr(rect_obj, 'x') and callable(rect_obj.x):
                                section_rects.append((rect_obj.x(), rect_obj.y(), rect_obj.width(), rect_obj.height()))
                            else:
                                print(f"[DEBUG] Invalid rect object in dictionary: {type(rect_obj)}")
                                section_rects.append((0, 0, 0, 0))
                        else:
                            # For other types, just use a placeholder tuple
                            print(f"[DEBUG] Using placeholder for rect object of type {type(rect)}")
                            section_rects.append((0, 0, 0, 0))
                    except Exception as e:
                        print(f"[DEBUG] Error processing rect in _get_extraction_state: {str(e)}")
                        # Add a placeholder tuple in case of error
                        section_rects.append((0, 0, 0, 0))
                regions_hash += (section, tuple(section_rects))
        except Exception as e:
            print(f"[DEBUG] Error in _get_extraction_state regions processing: {str(e)}")
            # Return empty tuple in case of error
            return (tuple(), tuple())

        # Convert column lines to a hashable format
        try:
            columns_hash = tuple()
            for section, lines in self.column_lines.items():
                try:
                    section_name = section.value if hasattr(section, 'value') else section
                    section_lines = []
                    for line in lines:
                        try:
                            if len(line) >= 2 and isinstance(line[0], QPoint):
                                # Format: (pair of QPoints with optional region index)
                                line_tuple = (line[0].x(), line[0].y(), line[1].x(), line[1].y())
                                if len(line) > 2:
                                    line_tuple += (line[2],)
                                section_lines.append(line_tuple)
                            else:
                                # Skip invalid format
                                print(f"[DEBUG] Skipping line with invalid format: {line}")
                                continue
                        except Exception as e:
                            print(f"[DEBUG] Error processing line in _get_extraction_state: {str(e)}")
                            # Add a placeholder tuple in case of error
                            section_lines.append((0, 0, 0, 0))
                    columns_hash += (section_name, tuple(section_lines))
                except Exception as e:
                    print(f"[DEBUG] Error processing section {section} in _get_extraction_state: {str(e)}")
                    # Skip this section in case of error
        except Exception as e:
            print(f"[DEBUG] Error in _get_extraction_state columns processing: {str(e)}")
            # Return partial result in case of error
            return (regions_hash, tuple())

        return (regions_hash, columns_hash)

    def retry_extraction(self):
        """Retry extraction with current parameters for the current page and show all pages data"""
        print(f"[DEBUG] Forcefully extracting regions from current page")

        # Clear extraction cache for all pages
        if hasattr(self, 'pdf_path') and self.pdf_path:
            clear_extraction_cache_for_pdf(self.pdf_path)
            print(f"[DEBUG] Cleared extraction cache for PDF: {self.pdf_path}")

        # Force extraction even if no changes were made
        self._last_extraction_state = None

        # For multi-page PDFs
        if self.multi_page_mode and hasattr(self, 'pdf_document') and len(self.pdf_document) > 1:
            print(f"[DEBUG] Multi-page mode detected, extracting data for current page {self.current_page_index + 1}")

            # Extract data for the current page using multipage_extract_page_data
            page_data = self.multipage_extract_page_data(self.current_page_index)

            # Store the extracted data in _all_pages_data
            if hasattr(self, '_all_pages_data') and self._all_pages_data:
                self._all_pages_data[self.current_page_index] = page_data
                print(f"[DEBUG] Updated _all_pages_data for page {self.current_page_index + 1}")

            # Now extract combined data from all pages to show in the JSON tree
            combined_data = self.extract_multi_page_invoice()

            # Update the cached extraction data with the combined data
            self._cached_extraction_data = combined_data

            # Update the JSON tree with the combined data from all pages
            self.update_json_tree(combined_data)

            # Show a brief notification
            QMessageBox.information(
                self,
                "Extraction Complete",
                f"Successfully extracted data from page {self.current_page_index + 1} and updated the view with data from all pages."
            )
        else:
            # For single-page PDFs, just force extraction for the current page
            self.update_extraction_results(force=True)

    def show_param_dialog(self):
        """Show dialog to adjust extraction parameters"""
        # Create a dialog to adjust extraction parameters
        dialog = QDialog(self)
        dialog.setWindowTitle("Adjust Extraction Parameters")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(650)  # Increased height for additional parameters

        # Create main layout
        layout = QVBoxLayout(dialog)

        # Add header text
        header_label = QLabel("Adjust parameters for PDF extraction")
        header_label.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(header_label)

        # Add explanation text
        explanation = QLabel("Modify these parameters to improve text extraction quality")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # Create tab widget for different sections
        tab_widget = QTabWidget()
        layout.addWidget(tab_widget)

        # Store widgets for each section
        section_params = {}

        # Create tabs for each section
        for section in ['header', 'items', 'summary']:
            # Create tab for section
            section_tab = QWidget()
            section_layout = QVBoxLayout(section_tab)

            # Create group for basic section parameters
            basic_group = QGroupBox(f"Basic {section.capitalize()} Parameters")
            basic_layout = QFormLayout(basic_group)

            # Row tolerance parameter
            row_tol = QSpinBox()
            row_tol.setRange(1, 999)  # Increased maximum value to 999 (effectively removing the constraint)
            row_tol.setValue(self.extraction_params.get(section, {}).get('row_tol', 10))
            row_tol.setToolTip("Tolerance for grouping text into rows (higher value = more text in same row)")
            basic_layout.addRow("Row Tolerance:", row_tol)

            # Add note about col_tol and min_rows
            note_label = QLabel("Note: 'Column Tolerance' and 'Minimum Rows' parameters can be added as custom parameters below if needed (only compatible with 'lattice' flavor).")
            note_label.setWordWrap(True)
            note_label.setStyleSheet("font-style: italic; color: #666; font-size: 9pt;")
            basic_layout.addRow("", note_label)

            # Add basic group to tab layout
            section_layout.addWidget(basic_group)

            # Create group for extraction options
            options_group = QGroupBox(f"{section.capitalize()} Extraction Options")
            options_layout = QFormLayout(options_group)

            # Flavor parameter
            flavor_combo = QComboBox()
            flavor_combo.addItems(['stream', 'lattice'])
            # Get flavor from section-specific params if available, otherwise from global params, with default 'stream'
            section_flavor = self.extraction_params.get(section, {}).get('flavor',
                            self.extraction_params.get('flavor', 'stream'))
            flavor_combo.setCurrentText(section_flavor)
            flavor_combo.setToolTip("'stream' is recommended for most documents, 'lattice' works better for tables with visible borders")
            options_layout.addRow("Extraction Flavor:", flavor_combo)

            # Split text parameter
            split_text = QCheckBox()
            # Get split_text from section-specific params if available, otherwise from global params, with default True
            section_split_text = self.extraction_params.get(section, {}).get('split_text',
                                self.extraction_params.get('split_text', True))
            split_text.setChecked(section_split_text)
            split_text.setToolTip("Split text that may contain multiple values")
            options_layout.addRow("Split Text:", split_text)

            # Strip text parameter
            strip_text = QLineEdit()
            # Get strip_text from section-specific params if available, otherwise from global params, with default '\n'
            section_strip_text = self.extraction_params.get(section, {}).get('strip_text',
                                self.extraction_params.get('strip_text', '\n'))
            strip_text.setText(section_strip_text)
            strip_text.setToolTip("Characters to strip from text (use \\n for newlines)")
            options_layout.addRow("Strip Text:", strip_text)

            # Removed parallel processing parameter as requested

            # Edge detection threshold
            edge_tol = QDoubleSpinBox()
            edge_tol.setRange(0.1, 10.0)
            edge_tol.setSingleStep(0.1)
            # Get edge_tol from section-specific params if available, otherwise from global params, with default 0.5
            section_edge_tol = self.extraction_params.get(section, {}).get('edge_tol',
                              self.extraction_params.get('edge_tol', 0.5))
            edge_tol.setValue(section_edge_tol)
            edge_tol.setToolTip("Threshold for edge detection in table extraction")
            options_layout.addRow("Edge Detection Threshold:", edge_tol)

            # Add options group to tab layout
            section_layout.addWidget(options_group)

            # Create group for custom parameters
            custom_group = QGroupBox(f"{section.capitalize()} Custom Parameters")
            custom_layout = QVBoxLayout(custom_group)

            # Add label for custom parameters
            custom_label = QLabel("Add up to 3 additional custom parameters:")
            custom_layout.addWidget(custom_label)

            # Create inputs for additional parameters
            additional_param_inputs = []

            for i in range(3):
                param_layout = QHBoxLayout()

                param_name_input = QLineEdit()
                param_name_input.setPlaceholderText(f"Parameter {i+1} name")
                param_name_input.setMaximumWidth(150)

                param_value_input = QLineEdit()
                param_value_input.setPlaceholderText(f"Parameter {i+1} value")

                # Pre-fill with existing custom parameters if available
                if f'custom_param_{i+1}_name' in self.extraction_params and f'custom_param_{i+1}_value' in self.extraction_params:
                    param_name_input.setText(self.extraction_params[f'custom_param_{i+1}_name'])
                    param_value_input.setText(str(self.extraction_params[f'custom_param_{i+1}_value']))

                param_layout.addWidget(param_name_input)
                param_layout.addWidget(param_value_input)

                additional_param_inputs.append((param_name_input, param_value_input))
                custom_layout.addLayout(param_layout)

            # Add custom group to tab layout
            section_layout.addWidget(custom_group)

            # Store widgets for later access
            section_params[section] = {
                'row_tol': row_tol,
                'flavor': flavor_combo,
                'split_text': split_text,
                'strip_text': strip_text,
                # Removed parallel parameter as requested
                'edge_tol': edge_tol,
                'additional_param_inputs': additional_param_inputs
            }

            # Add tab to tab widget
            tab_widget.addTab(section_tab, f"{section.capitalize()}")

        # Add a note about section-specific parameters
        note_label = QLabel("Note: Each section tab contains all parameters for that section, including extraction options and custom parameters.")
        note_label.setWordWrap(True)
        note_label.setStyleSheet("font-style: italic; color: #666;")
        layout.addWidget(note_label)

        # Add buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        ok_button = button_box.button(QDialogButtonBox.Ok)
        ok_button.setText("Extract")
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        # Show dialog and process result
        if dialog.exec_() == QDialog.Accepted:
            # Update extraction parameters for each section
            for section, widgets in section_params.items():
                if section not in self.extraction_params:
                    self.extraction_params[section] = {}

                # Update basic parameters
                self.extraction_params[section]['row_tol'] = widgets['row_tol'].value()

                # Remove col_tol and min_rows if they exist (they're now custom parameters)
                if 'col_tol' in self.extraction_params[section]:
                    del self.extraction_params[section]['col_tol']
                if 'min_rows' in self.extraction_params[section]:
                    del self.extraction_params[section]['min_rows']

                # Update section-specific extraction options
                self.extraction_params[section]['flavor'] = widgets['flavor'].currentText()
                self.extraction_params[section]['split_text'] = widgets['split_text'].isChecked()
                self.extraction_params[section]['strip_text'] = widgets['strip_text'].text()
                # Removed parallel parameter as requested
                self.extraction_params[section]['edge_tol'] = widgets['edge_tol'].value()

                # Update custom parameters for this section
                for i, (name_input, value_input) in enumerate(widgets['additional_param_inputs']):
                    param_name = name_input.text().strip()
                    param_value = value_input.text().strip()

                    # Clear previous custom parameters for this section
                    param_name_key = f'{section}_custom_param_{i+1}_name'
                    param_value_key = f'{section}_custom_param_{i+1}_value'

                    if param_name_key in self.extraction_params:
                        del self.extraction_params[param_name_key]
                    if param_value_key in self.extraction_params:
                        del self.extraction_params[param_value_key]

                    # Add new custom parameters if provided
                    if param_name and param_value:
                        self.extraction_params[param_name_key] = param_name

                        # Try to convert value to appropriate type
                        try:
                            if param_value.lower() in ['true', 'false']:
                                self.extraction_params[param_value_key] = param_value.lower() == 'true'
                            elif param_value.isdigit():
                                self.extraction_params[param_value_key] = int(param_value)
                            elif param_value.replace('.', '', 1).isdigit():
                                self.extraction_params[param_value_key] = float(param_value)
                            else:
                                self.extraction_params[param_value_key] = param_value
                        except ValueError:
                            self.extraction_params[param_value_key] = param_value

                        print(f"[DEBUG] Added custom parameter for {section}: {param_name} = {param_value}")

            # Get the currently active tab to determine which section was modified
            active_tab_index = tab_widget.currentIndex()
            active_section = ['header', 'items', 'summary'][active_tab_index]

            print(f"[DEBUG] Active tab is {active_tab_index}, corresponding to section: {active_section}")

            # Only extract data for the section that was modified
            if self.regions[active_section]:  # Only extract if regions exist for this section
                print(f"[DEBUG] Extracting data for modified section: {active_section}")
                self.extract_with_new_params(active_section)
            else:
                print(f"[DEBUG] No regions defined for section: {active_section}")
                QMessageBox.warning(self, "No Regions", f"No regions defined for {active_section} section.")

    def extract_with_new_params(self, section):
        """Extract data with new parameters for the specified section

        Args:
            section (str): The section to extract ('header', 'items', or 'summary')
        """
        try:
            # Get regions for this section based on multi-page mode
            if self.multi_page_mode:
                if self.current_page_index in self.page_regions and section in self.page_regions[self.current_page_index]:
                    section_regions = self.page_regions[self.current_page_index][section]
                else:
                    print(f"[DEBUG] No regions found for {section} in multi-page mode")
                    section_regions = []
            else:
                section_regions = self.regions.get(section, [])

            if not section_regions:
                QMessageBox.warning(self, "No Regions", f"No regions defined for {section} section.")
                return

            # Get column lines for this section based on multi-page mode
            section_enum = RegionType(section)

            if self.multi_page_mode:
                # For multi-page mode, get column lines from page_column_lines
                if self.current_page_index in self.page_column_lines:
                    if section_enum in self.page_column_lines[self.current_page_index]:
                        section_column_lines = self.page_column_lines[self.current_page_index][section_enum]
                    elif section in self.page_column_lines[self.current_page_index]:
                        section_column_lines = self.page_column_lines[self.current_page_index][section]
                    else:
                        print(f"[DEBUG] No column lines found for {section} in multi-page mode")
                        section_column_lines = []
                else:
                    print(f"[DEBUG] No column lines found for page {self.current_page_index + 1}")
                    section_column_lines = []
            else:
                # For single-page mode, get column lines from column_lines
                if section_enum in self.column_lines:
                    section_column_lines = self.column_lines[section_enum]
                elif section in self.column_lines:
                    section_column_lines = self.column_lines[section]
                else:
                    print(f"[DEBUG] No column lines found for {section} in single-page mode")
                    section_column_lines = []

            # Get PDF page dimensions and scale factors
            scale_factors = get_scale_factors(self.pdf_path, self.current_page_index)
            scale_x = scale_factors['scale_x']
            scale_y = scale_factors['scale_y']
            page_height = scale_factors['page_height']

            # Convert regions to table areas
            table_areas = []
            for region_item in section_regions:
                # Handle both QRect objects and dictionary format
                rect, _ = extract_rect_and_label(region_item)

                # Skip invalid regions
                if rect is None:
                    log_warning(f"Skipping invalid region in {section}")
                    continue

                # Convert QRect to PDF coordinates (bottom-left coordinate system)
                x1 = rect.x() * scale_x
                y1 = page_height - (rect.y() * scale_y)  # Convert top y1 to bottom y1
                x2 = (rect.x() + rect.width()) * scale_x
                y2 = page_height - ((rect.y() + rect.height()) * scale_y)  # Convert bottom y2 to bottom y2

                # Add to table areas
                table_areas.append([x1, y1, x2, y2])

            # Convert column lines to columns list
            columns_list = []
            for line in section_column_lines:
                # Check if line is in the new format (pair of QPoints with optional region index)
                if len(line) >= 2 and isinstance(line[0], QPoint):
                    # Get the x-coordinate and convert to PDF coordinates
                    x = line[0].x() * scale_x
                    columns_list.append(x)

            # Sort columns by x-coordinate
            columns_list.sort()

            # Get section-specific parameters
            section_params = self.extraction_params.get(section, {})

            # Prepare extraction parameters - include all necessary parameters
            extraction_params = {
                'pages': str(self.current_page_index + 1),  # 1-based page number
            }

            # Add all standard parameters from section_params
            for param_name in ['split_text', 'strip_text', 'edge_tol', 'row_tol']:
                if param_name in section_params:
                    extraction_params[param_name] = section_params[param_name]

            # Don't pass table_areas and columns directly to extraction_params
            # They will be properly formatted by extract_table/extract_tables functions

            # Determine the flavor - default to 'stream' if not specified
            flavor = section_params.get('flavor', 'stream')

            # Check if we have columns and need to adjust the flavor
            if columns_list and flavor == 'lattice':
                print(f"[WARNING] Columns are defined but flavor is set to 'lattice'. Changing to 'stream' flavor.")
                flavor = 'stream'

            # Set the flavor parameter
            extraction_params['flavor'] = flavor

            # Add flavor-specific parameters
            if flavor == 'lattice':
                # These parameters are only compatible with lattice flavor
                for param_name in ['col_tol', 'min_rows']:
                    if param_name in section_params:
                        extraction_params[param_name] = section_params[param_name]

            # Add any custom parameters for this section
            for i in range(1, 10):  # Support up to 9 custom parameters
                param_name_key = f'{section}_custom_param_{i}_name'
                param_value_key = f'{section}_custom_param_{i}_value'

                if param_name_key in self.extraction_params and param_value_key in self.extraction_params:
                    param_name = self.extraction_params[param_name_key]
                    param_value = self.extraction_params[param_value_key]
                    extraction_params[param_name] = param_value

            # Also check for global custom parameters
            for key, value in self.extraction_params.items():
                if key.startswith('custom_param_') and not key.startswith(f'{section}_custom_param_'):
                    # This is a global custom parameter, extract the actual parameter name
                    if '_name' in key and key.replace('_name', '_value') in self.extraction_params:
                        param_name = value
                        param_value = self.extraction_params[key.replace('_name', '_value')]
                        extraction_params[param_name] = param_value

            print(f"[DEBUG] Extracting {section} with parameters: {extraction_params}")

            # Save the current extraction data for other sections
            current_data = self._cached_extraction_data.copy() if hasattr(self, '_cached_extraction_data') else {}

            # Check if we have multiple table areas
            if len(table_areas) > 1:
                # Extract multiple tables using pdf_extraction_utils
                print(f"[DEBUG] Extracting {len(table_areas)} tables for {section} section")
                print(f"[DEBUG] Ignoring cached data and forcing re-extraction due to parameter adjustment")

                # Clear the extraction cache for this section
                clear_extraction_cache_for_section(self.pdf_path, self.current_page_index + 1, section)

                # Use multi-method extraction
                extraction_method = getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
                table_dfs = extract_with_method(
                    pdf_path=self.pdf_path,
                    extraction_method=extraction_method,
                    page_number=self.current_page_index + 1,  # 1-based page number
                    table_areas=table_areas,
                    columns_list=[columns_list] if columns_list else None,
                    section_type=section,
                    extraction_params=self.extraction_params,  # Use the actual extraction_params from the class
                    use_cache=False  # Don't use cache for parameter testing
                )

                if table_dfs and len(table_dfs) > 0:
                    # Create a new cached extraction data dictionary with the current data
                    self._cached_extraction_data = {
                        'header': current_data.get('header', []),
                        'items': current_data.get('items', []),
                        'summary': current_data.get('summary', [])
                    }

                    # Update only the section that was modified
                    self._cached_extraction_data[section] = table_dfs
                    print(f"[DEBUG] Updated cached extraction data for {section} with {len(table_dfs)} tables")

                    # Update the JSON tree
                    self.update_json_tree(self._cached_extraction_data)

                    # If we're in multi-page mode, update the stored data for the current page
                    if self.multi_page_mode and hasattr(self, '_all_pages_data') and self._all_pages_data:
                        if self.current_page_index < len(self._all_pages_data):
                            # Update the data for the current page with the new extraction results
                            self._all_pages_data[self.current_page_index] = self._cached_extraction_data.copy()
                            print(f"[DEBUG] Updated stored extraction data for page {self.current_page_index + 1} after parameter adjustment (multiple tables)")

                            # Make sure to update the page_regions and page_column_lines dictionaries
                            self.page_regions[self.current_page_index] = self.regions.copy()
                            self.page_column_lines[self.current_page_index] = self.column_lines.copy()
                            print(f"[DEBUG] Updated page_regions and page_column_lines for page {self.current_page_index + 1} after parameter adjustment (multiple tables)")

                    # Force a re-extraction to update the display
                    self._last_extraction_state = None

                    # Show success message
                    QMessageBox.information(
                        self,
                        f"{section.capitalize()} Extraction Results",
                        f"Successfully extracted {len(table_dfs)} tables for {section} section."
                    )
                    return
                else:
                    print(f"[DEBUG] No valid tables found for {section} section")
                    QMessageBox.warning(
                        self,
                        f"{section.capitalize()} Extraction Results",
                        f"No valid tables found for {section} section."
                    )
                    return
            else:
                # Extract single table using pdf_extraction_utils
                print(f"[DEBUG] Ignoring cached data and forcing re-extraction due to parameter adjustment")

                # Clear the extraction cache for this section
                clear_extraction_cache_for_section(self.pdf_path, self.current_page_index + 1, section)

                # Use multi-method extraction
                extraction_method = getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
                table_df = extract_with_method(
                    pdf_path=self.pdf_path,
                    extraction_method=extraction_method,
                    page_number=self.current_page_index + 1,  # 1-based page number
                    table_areas=[table_areas[0]] if table_areas else None,  # First table area
                    columns_list=[columns_list] if columns_list else None,
                    section_type=section,
                    extraction_params=self.extraction_params,  # Use the actual extraction_params from the class
                    use_cache=False  # Don't use cache for parameter testing
                )

            if table_df is not None and not table_df.empty:
                print(f"[DEBUG] Successfully extracted {section} table with {len(table_df)} rows and {len(table_df.columns)} columns")

                # Create a new cached extraction data dictionary with the current data
                self._cached_extraction_data = {
                    'header': current_data.get('header', []),
                    'items': current_data.get('items', []),
                    'summary': current_data.get('summary', [])
                }

                # Update only the section that was modified
                self._cached_extraction_data[section] = table_df
                print(f"[DEBUG] Updated cached extraction data for {section} with table of {len(table_df)} rows")

                # Update the JSON tree
                self.update_json_tree(self._cached_extraction_data)

                # If we're in multi-page mode, update the stored data for the current page
                if self.multi_page_mode and hasattr(self, '_all_pages_data') and self._all_pages_data:
                    if self.current_page_index < len(self._all_pages_data):
                        # Update the data for the current page with the new extraction results
                        self._all_pages_data[self.current_page_index] = self._cached_extraction_data.copy()
                        print(f"[DEBUG] Updated stored extraction data for page {self.current_page_index + 1} after parameter adjustment")

                        # Make sure to update the page_regions and page_column_lines dictionaries
                        self.page_regions[self.current_page_index] = self.regions.copy()
                        self.page_column_lines[self.current_page_index] = self.column_lines.copy()
                        print(f"[DEBUG] Updated page_regions and page_column_lines for page {self.current_page_index + 1} after parameter adjustment")

                # Force a re-extraction to update the display
                self._last_extraction_state = None

                # Show success message
                QMessageBox.information(
                    self,
                    f"{section.capitalize()} Extraction Results",
                    f"Successfully extracted {section} table with {len(table_df)} rows and {len(table_df.columns)} columns."
                )
            else:
                print(f"[DEBUG] No data extracted for {section} section")
                QMessageBox.warning(
                    self,
                    f"{section.capitalize()} Extraction Results",
                    f"No data extracted for {section} section."
                )

        except Exception as e:
            print(f"[ERROR] Failed to extract with new parameters: {str(e)}")
            import traceback
            traceback.print_exc()

            QMessageBox.critical(
                self,
                "Extraction Error",
                f"An error occurred during extraction: {str(e)}"
            )

    def extract_page_data(self, page_index):
        """Extract data from a specific page

        Args:
            page_index (int): The index of the PDF page (0-based)

        Returns:
            tuple: (header_df, items_df, summary_df) - DataFrames containing the extracted data
        """
        if not self.pdf_document:
            return None, None, None

        try:
            print(f"\n[DEBUG] extract_page_data called for page {page_index + 1}")
            print(f"[DEBUG] Current page index: {self.current_page_index + 1}")
            print(f"[DEBUG] PDF has {len(self.pdf_document)} pages")

            # Save current page index
            current_page = self.current_page_index

            # Switch to the requested page if needed
            if current_page != page_index:
                print(f"[DEBUG] Switching to page {page_index + 1} for extraction")
                self.current_page_index = page_index
                self.display_current_page()
            # Check the source of the PDF loading
            pdf_source = getattr(self, 'pdf_loading_source', 'direct')
            print(f"[DEBUG] PDF loading source: {pdf_source}")

            # Initialize extraction parameters based on the source
            if pdf_source == "template_manager":
                # For template_manager source, use the extraction parameters from the template
                # These should already be set in apply_template method
                if not hasattr(self, 'extraction_params') or not self.extraction_params:
                    print(f"[WARNING] No extraction parameters found for template_manager source, using defaults")
                    self.extraction_params = {
                        'header': {'row_tol': 5},
                        'items': {'row_tol': 15},
                        'summary': {'row_tol': 10},
                        'flavor': 'stream',
                        'strip_text': '\n'
                    }
            else:
                # For direct loading, use default parameters
                if not hasattr(self, 'extraction_params') or not self.extraction_params:
                    self.extraction_params = {
                        'header': {'row_tol': 5},
                        'items': {'row_tol': 15},
                        'summary': {'row_tol': 10},
                        'flavor': 'stream',
                        'strip_text': '\n'
                    }

            # Ensure extraction parameters have the required structure
            if not isinstance(self.extraction_params, dict):
                print(f"[WARNING] Extraction parameters are not a dictionary: {type(self.extraction_params)}")
                self.extraction_params = {
                    'header': {'row_tol': 5},
                    'items': {'row_tol': 15},
                    'summary': {'row_tol': 10},
                    'flavor': 'stream',
                    'strip_text': '\n'
                }

            # Ensure all section parameters are properly initialized
            for section in ['header', 'items', 'summary']:
                if section not in self.extraction_params:
                    self.extraction_params[section] = {}

                # Ensure section parameters are dictionaries
                if not isinstance(self.extraction_params[section], dict):
                    self.extraction_params[section] = {}

            # Get global parameters to use as defaults for section-specific parameters
            global_flavor = self.extraction_params.get('flavor', 'stream')
            global_split_text = self.extraction_params.get('split_text', True)
            global_strip_text = self.extraction_params.get('strip_text', '\n')

            # Ensure each section has the necessary parameters
            for section in ['header', 'items', 'summary']:
                section_params = self.extraction_params[section]

                # Add missing parameters to section if they don't exist
                if 'row_tol' not in section_params:
                    if section == 'header':
                        section_params['row_tol'] = 5
                    elif section == 'items':
                        section_params['row_tol'] = 15
                    else:  # summary
                        section_params['row_tol'] = 10

                if 'flavor' not in section_params:
                    section_params['flavor'] = global_flavor

                if 'split_text' not in section_params:
                    section_params['split_text'] = global_split_text

                if 'strip_text' not in section_params:
                    section_params['strip_text'] = global_strip_text

                # Add edge_tol if not present
                if 'edge_tol' not in section_params:
                    section_params['edge_tol'] = 0.5

            # Ensure flavor is set at the global level
            if 'flavor' not in self.extraction_params:
                self.extraction_params['flavor'] = 'stream'

            # Look for custom parameters in the extraction_params
            for key, value in list(self.extraction_params.items()):
                if key.startswith('custom_param_'):
                    print(f"[DEBUG] Found custom parameter: {key} = {value}")

            # Print detailed extraction parameters structure for debugging
            print(f"[DEBUG] ===== EXTRACTION PARAMETERS STRUCTURE =====")
            print(f"[DEBUG] extraction_params type: {type(self.extraction_params).__name__}")
            print(f"[DEBUG] extraction_params keys: {list(self.extraction_params.keys())}")

            # Print section-specific parameters
            for section in ['header', 'items', 'summary']:
                section_params = self.extraction_params.get(section, {})
                print(f"[DEBUG] {section.capitalize()} extraction parameters: {section_params}")
                print(f"[DEBUG] {section.capitalize()} parameters type: {type(section_params).__name__}")
                if isinstance(section_params, dict):
                    print(f"[DEBUG] {section.capitalize()} parameters keys: {list(section_params.keys())}")
                    for key, value in section_params.items():
                        print(f"[DEBUG]   - {section}.{key}: {value} (type: {type(value).__name__})")

            # Print global extraction parameters
            print(f"[DEBUG] Global extraction parameters:")
            for key, value in self.extraction_params.items():
                if key not in ['header', 'items', 'summary']:
                    print(f"[DEBUG]   - {key}: {value} (type: {type(value).__name__})")

            print(f"[DEBUG] ===== END EXTRACTION PARAMETERS STRUCTURE =====")

            # Print extraction parameters for debugging
            print(f"[DEBUG] Using extraction parameters: {self.extraction_params}")
            for section in ['header', 'items', 'summary']:
                print(f"[DEBUG] {section.capitalize()} extraction parameters: {self.extraction_params.get(section, {})}")

            for key, value in self.extraction_params.items():
                if key not in ['header', 'items', 'summary']:
                    print(f"[DEBUG] Global parameter '{key}': {value}")

            # Determine which template page to use based on settings
            template_page_index = self.get_template_page_index(page_index)
            print(f"[DEBUG] In extract_page_data: Using template page {template_page_index + 1 if template_page_index >= 0 else 'combined'} for PDF page {page_index + 1}")

            # Get regions and column lines for this page
            current_regions = {}
            current_column_lines = {}

            # Special case for combined template pages
            if template_page_index == -1 and hasattr(self, 'use_middle_page') and self.use_middle_page:
                print(f"[DEBUG] Special case: 1-page PDF with use_middle_page=True")
                print(f"[DEBUG] Combining first and last template pages for PDF page {page_index + 1}")

                # Get the first and last template page indices
                first_template_index = 0
                last_template_index = min(2, len(self.page_configs) - 1) if len(self.page_configs) > 2 else len(self.page_configs) - 1

                # Combine regions and column lines from first and last template pages
                combined_regions, combined_column_lines = self.combine_template_pages(first_template_index, last_template_index)
                current_regions = combined_regions
                current_column_lines = combined_column_lines
            else:
                # Get regions for this page
                if self.multi_page_mode:
                    if isinstance(self.page_regions, dict) and page_index in self.page_regions:
                        current_regions = self.page_regions[page_index]
                    elif isinstance(self.page_regions, list) and template_page_index < len(self.page_regions):
                        current_regions = self.page_regions[template_page_index]
                    else:
                        current_regions = {'header': [], 'items': [], 'summary': []}

                    # Get column lines for this page
                    if isinstance(self.page_column_lines, dict) and page_index in self.page_column_lines:
                        current_column_lines = self.page_column_lines[page_index]
                        print(f"[DEBUG] Using column lines from page_column_lines[{page_index}]")

                        # Debug column lines
                        for section, lines in current_column_lines.items():
                            section_name = section.value if hasattr(section, 'value') else section
                            print(f"[DEBUG] Found {len(lines)} column lines for {section_name}")

                    elif isinstance(self.page_column_lines, list) and template_page_index < len(self.page_column_lines):
                        current_column_lines = self.page_column_lines[template_page_index]
                        print(f"[DEBUG] Using column lines from page_column_lines[{template_page_index}] (template)")
                    else:
                        current_column_lines = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}
                        print(f"[DEBUG] No column lines found for page {page_index}, using empty dictionary")

                    # Allow summary section on all pages
                    # The previous code removed summary sections from all but the last page
                    # We're now keeping summary sections on all pages
                    if len(self.pdf_document) > 1:
                        print(f"[DEBUG] Keeping summary section for page {page_index + 1} in multi-page PDF")
                else:
                    # Single-page mode
                    current_regions = self.regions
                    current_column_lines = self.column_lines

            # Get PDF page dimensions and scale factors
            scale_factors = get_scale_factors(self.pdf_path, page_index)
            scale_x = scale_factors['scale_x']
            scale_y = scale_factors['scale_y']
            page_height = scale_factors['page_height']

            print(f"[DEBUG] PDF dimensions: {scale_factors['page_width']} x {page_height}")
            print(f"[DEBUG] Scale factors: X={scale_x}, Y={scale_y}")

            # Convert regions to table areas
            table_areas = {}
            for section, rects in current_regions.items():
                table_areas[section] = []
                for region_item in rects:
                    # Handle both QRect objects and dictionary format
                    rect, region_label = extract_rect_and_label(region_item)

                    # Skip invalid regions
                    if rect is None:
                        log_warning(f"Skipping invalid region in {section}")
                        continue

                    # Convert QRect to PDF coordinates (bottom-left coordinate system)
                    x1 = rect.x() * scale_x
                    y1 = page_height - (rect.y() * scale_y)  # Convert top y1 to bottom y1
                    x2 = (rect.x() + rect.width()) * scale_x
                    y2 = page_height - ((rect.y() + rect.height()) * scale_y)  # Convert bottom y2 to bottom y2

                    # Add to table areas
                    table_areas[section].append([x1, y1, x2, y2])
                    print(f"[DEBUG] Converted {section} region: Display({rect.x()},{rect.y()},{rect.x()+rect.width()},{rect.y()+rect.height()}) -> PDF({x1},{y1},{x2},{y2})")

            # Convert column lines to columns list
            columns_list = {}
            for section, lines in current_column_lines.items():
                section_name = section.value if hasattr(section, 'value') else section
                columns_list[section_name] = []

                # Process each region separately
                for region_idx, region in enumerate(current_regions.get(section_name, [])):
                    region_columns = []

                    # Find column lines for this region
                    for line in lines:
                        # Check if line is in the new format (pair of QPoints with optional region index)
                        if len(line) >= 2 and isinstance(line[0], QPoint):
                            # Check if this line belongs to the current region
                            if len(line) > 2 and isinstance(line[2], int):
                                # Line has region index
                                if line[2] == region_idx:
                                    # Line matches current region
                                    x = line[0].x() * scale_x
                                    region_columns.append(x)
                                    print(f"[DEBUG] Found column for region {region_idx} at x={line[0].x()} -> PDF({x}) (matched region index)")
                            else:
                                # Line has no region index, associate with all regions in multipage mode
                                # or only with first region in single-page mode
                                if self.multi_page_mode or region_idx == 0:
                                    x = line[0].x() * scale_x
                                    region_columns.append(x)
                                    print(f"[DEBUG] Found column at x={line[0].x()} -> PDF({x}) for region {region_idx} (no region index)")
                        else:
                            # Skip invalid format
                            print(f"[DEBUG] Skipping column line with invalid format: {line}")

                    # Sort columns by x-coordinate
                    region_columns.sort()

                    # Add columns for this region
                    if region_columns:
                        # Convert to comma-separated string
                        col_str = ','.join(map(str, region_columns))
                        columns_list.setdefault(section_name, []).append(col_str)
                        print(f"[DEBUG] Added columns for {section_name} region {region_idx}: {col_str}")
                    else:
                        # Empty string for regions with no column lines
                        columns_list.setdefault(section_name, []).append('')
                        print(f"[DEBUG] No columns for {section_name} region {region_idx}")

            # Extract data for each section
            header_df = None
            items_df = None
            summary_df = None

            # Extract header data
            if 'header' in table_areas and table_areas['header']:
                header_tables = []
                for i, area in enumerate(table_areas['header']):
                    # Get columns for this specific region
                    columns = None
                    if 'header' in columns_list and i < len(columns_list['header']):
                        columns = columns_list['header'][i]
                        print(f"[DEBUG] Using columns for header region {i}: {columns}")
                        print(f"[DEBUG] Columns type: {type(columns)}")
                    # Get params without defaults
                    params = self.extraction_params.get('header', {})

                    # Create a minimal set of parameters without defaults
                    extraction_params = {
                        'pages': str(page_index + 1),  # 1-based page number
                        'header': {'pages': str(page_index + 1)},
                        'items': {'pages': str(page_index + 1)},
                        'summary': {'pages': str(page_index + 1)}
                    }

                    # Determine the flavor - default to 'stream' if not specified
                    flavor = params.get('flavor', 'stream')

                    # Check if we have columns and need to adjust the flavor
                    if columns and flavor == 'lattice':
                        print(f"[WARNING] Columns are defined but flavor is set to 'lattice'. Changing to 'stream' flavor.")
                        flavor = 'stream'

                    # Set the flavor parameter
                    extraction_params['flavor'] = flavor

                    # Only add other parameters that exist in params
                    for key, value in params.items():
                        if key != 'flavor':  # Skip flavor as we've already handled it
                            extraction_params[key] = value

                    print(f"[DEBUG] extract_table called for header on page {page_index + 1}")
                    print(f"[DEBUG] PDF path: {self.pdf_path}")
                    print(f"[DEBUG] Table area: {area}")
                    print(f"[DEBUG] Columns: {columns}")
                    print(f"[DEBUG] Using extraction params: {extraction_params}")

                    # Prepare additional parameters from custom parameters
                    additional_params = {}
                    for i in range(1, 4):  # Check for up to 3 custom parameters
                        param_name_key = f'custom_param_{i}_name'
                        param_value_key = f'custom_param_{i}_value'

                        if param_name_key in self.extraction_params and param_value_key in self.extraction_params:
                            param_name = self.extraction_params[param_name_key]
                            param_value = self.extraction_params[param_value_key]

                            if param_name and param_value is not None:
                                additional_params[param_name] = param_value
                                print(f"[DEBUG] Using custom parameter for header: {param_name} = {param_value}")

                    # Print extraction parameters before calling extract_table
                    print(f"[DEBUG] Calling extract_table for header with extraction_params: {self.extraction_params}")
                    print(f"[DEBUG] Header section parameters: {self.extraction_params.get('header', {})}")

                    # Extract table using multi-method extraction
                    extraction_method = getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
                    df = extract_with_method(
                        pdf_path=self.pdf_path,
                        extraction_method=extraction_method,
                        page_number=page_index + 1,  # 1-based page number
                        table_areas=[area],
                        columns_list=[columns] if columns else None,
                        section_type='header',
                        extraction_params=self.extraction_params,  # Use the actual extraction_params from the class
                        use_cache=False  # Disable caching to ensure fresh extraction with column lines
                    )

                    if df is not None and not df.empty:
                        header_tables.append(df)

                # Combine header tables
                if header_tables:
                    # Always return a list for consistency across all pages
                    # This ensures that page 1 and page 2 both have the same data structure
                    header_df = []

                    # Process each table
                    for i, df in enumerate(header_tables):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            # Add page number columns
                            df['page_number'] = page_index + 1  # 1-based page number
                            df['_page_number'] = page_index + 1  # Internal tracking

                            # Add region labels if they don't exist
                            if 'region_label' not in df.columns:
                                # Check if region labels have already been set for this page and section
                                if (page_index in self._region_labels_set and
                                    'header' in self._region_labels_set[page_index] and
                                    i in self._region_labels_set[page_index]['header']):
                                    print(f"[DEBUG] Region labels already set for header DataFrame {i+1} on page {page_index+1}, skipping")
                                    continue

                                # Use the actual region label from the region dictionary
                                region_label = f"H{i+1}"  # Default
                                if 'header' in self.page_regions.get(page_index, {}) and i < len(self.page_regions[page_index]['header']):
                                    # Get the label from the corresponding header region
                                    region_item = self.page_regions[page_index]['header'][i]
                                    if isinstance(region_item, dict) and 'label' in region_item:
                                        region_label = region_item['label']

                                # Create region labels with page number to ensure uniqueness
                                # IMPORTANT: Preserve the original region label (H1, H2, etc.) exactly as it is
                                df['region_label'] = [f"{region_label}_R{j+1}_P{page_index+1}"
                                                    for j in range(len(df))]
                                print(f"[DEBUG] Created region labels for header DataFrame {i+1}: {df['region_label'].tolist()}")
                                print(f"[DEBUG] Using original region label: {region_label} (preserving region number)")

                                # Mark these region labels as set
                                if page_index not in self._region_labels_set:
                                    self._region_labels_set[page_index] = {}
                                if 'header' not in self._region_labels_set[page_index]:
                                    self._region_labels_set[page_index]['header'] = {}
                                self._region_labels_set[page_index]['header'][i] = True
                            else:
                                # DO NOT modify region labels - preserve them exactly as they are
                                # Just verify that the labels are present and log them
                                labels = df['region_label'].tolist()
                                print(f"[DEBUG] Preserving original region labels for header DataFrame {i+1}: {labels}")
                                # No modification to the labels

                                # Mark these region labels as set
                                if page_index not in self._region_labels_set:
                                    self._region_labels_set[page_index] = {}
                                if 'header' not in self._region_labels_set[page_index]:
                                    self._region_labels_set[page_index]['header'] = {}
                                self._region_labels_set[page_index]['header'][i] = True
                            print(f"[DEBUG] Added page number {page_index+1} to header DataFrame {i+1}")
                            header_df.append(df)

                    print(f"[DEBUG] Returning header data as a list with {len(header_df)} DataFrames for consistency")

            # Extract items data
            if 'items' in table_areas and table_areas['items']:
                items_tables = []
                for i, area in enumerate(table_areas['items']):
                    # Get columns for this specific region
                    columns = None
                    if 'items' in columns_list and i < len(columns_list['items']):
                        columns = columns_list['items'][i]
                        print(f"[DEBUG] Using columns for items region {i}: {columns}")
                        print(f"[DEBUG] Columns type: {type(columns)}")
                    # Get params without defaults
                    params = self.extraction_params.get('items', {})

                    # Create a minimal set of parameters without defaults
                    extraction_params = {
                        'pages': str(page_index + 1),  # 1-based page number
                        'header': {'pages': str(page_index + 1)},
                        'items': {'pages': str(page_index + 1)},
                        'summary': {'pages': str(page_index + 1)}
                    }

                    # Determine the flavor - default to 'stream' if not specified
                    flavor = params.get('flavor', 'stream')

                    # Check if we have columns and need to adjust the flavor
                    if columns and flavor == 'lattice':
                        print(f"[WARNING] Columns are defined but flavor is set to 'lattice'. Changing to 'stream' flavor.")
                        flavor = 'stream'

                    # Set the flavor parameter
                    extraction_params['flavor'] = flavor

                    # Only add other parameters that exist in params
                    for key, value in params.items():
                        if key != 'flavor':  # Skip flavor as we've already handled it
                            extraction_params[key] = value

                    print(f"[DEBUG] extract_table called for items on page {page_index + 1}")
                    print(f"[DEBUG] PDF path: {self.pdf_path}")
                    print(f"[DEBUG] Table area: {area}")
                    print(f"[DEBUG] Columns: {columns}")
                    print(f"[DEBUG] Using extraction params: {extraction_params}")

                    # Prepare additional parameters from custom parameters
                    additional_params = {}
                    for i in range(1, 4):  # Check for up to 3 custom parameters
                        param_name_key = f'custom_param_{i}_name'
                        param_value_key = f'custom_param_{i}_value'

                        if param_name_key in self.extraction_params and param_value_key in self.extraction_params:
                            param_name = self.extraction_params[param_name_key]
                            param_value = self.extraction_params[param_value_key]

                            if param_name and param_value is not None:
                                additional_params[param_name] = param_value
                                print(f"[DEBUG] Using custom parameter for items: {param_name} = {param_value}")

                    # Print extraction parameters before calling extract_table
                    print(f"[DEBUG] Calling extract_table for items with extraction_params: {self.extraction_params}")
                    print(f"[DEBUG] Items section parameters: {self.extraction_params.get('items', {})}")

                    # Extract table using multi-method extraction
                    extraction_method = getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
                    df = extract_with_method(
                        pdf_path=self.pdf_path,
                        extraction_method=extraction_method,
                        page_number=page_index + 1,  # 1-based page number
                        table_areas=[area],
                        columns_list=[columns] if columns else None,
                        section_type='items',
                        extraction_params=self.extraction_params,  # Use the actual extraction_params from the class
                        use_cache=False  # Disable caching to ensure fresh extraction with column lines
                    )

                    if df is not None and not df.empty:
                        items_tables.append(df)

                # Combine items tables
                if items_tables:
                    # Always return a list for consistency across all pages
                    # This ensures that page 1 and page 2 both have the same data structure
                    items_df = []

                    # Process each table
                    for i, df in enumerate(items_tables):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            # Add page number columns
                            df['page_number'] = page_index + 1  # 1-based page number
                            df['_page_number'] = page_index + 1  # Internal tracking

                            # Add region labels if they don't exist
                            if 'region_label' not in df.columns:
                                # Check if region labels have already been set for this page and section
                                if (page_index in self._region_labels_set and
                                    'items' in self._region_labels_set[page_index] and
                                    i in self._region_labels_set[page_index]['items']):
                                    print(f"[DEBUG] Region labels already set for items DataFrame {i+1} on page {page_index+1}, skipping")
                                    continue

                                # Use the actual region label from the region dictionary
                                region_label = f"I{i+1}"  # Default
                                if 'items' in self.page_regions.get(page_index, {}) and i < len(self.page_regions[page_index]['items']):
                                    # Get the label from the corresponding items region
                                    region_item = self.page_regions[page_index]['items'][i]
                                    if isinstance(region_item, dict) and 'label' in region_item:
                                        region_label = region_item['label']

                                # Create region labels with page number to ensure uniqueness
                                # IMPORTANT: Preserve the original region label (I1, I2, etc.) exactly as it is
                                df['region_label'] = [f"{region_label}_R{j+1}_P{page_index+1}"
                                                    for j in range(len(df))]
                                print(f"[DEBUG] Created region labels for items DataFrame {i+1}: {df['region_label'].tolist()}")
                                print(f"[DEBUG] Using original region label: {region_label} (preserving region number)")

                                # Mark these region labels as set
                                if page_index not in self._region_labels_set:
                                    self._region_labels_set[page_index] = {}
                                if 'items' not in self._region_labels_set[page_index]:
                                    self._region_labels_set[page_index]['items'] = {}
                                self._region_labels_set[page_index]['items'][i] = True
                            else:
                                # DO NOT modify region labels - preserve them exactly as they are
                                # Just verify that the labels are present and log them
                                labels = df['region_label'].tolist()
                                print(f"[DEBUG] Preserving original region labels for items DataFrame {i+1}: {labels}")
                                # No modification to the labels

                                # Mark these region labels as set
                                if page_index not in self._region_labels_set:
                                    self._region_labels_set[page_index] = {}
                                if 'items' not in self._region_labels_set[page_index]:
                                    self._region_labels_set[page_index]['items'] = {}
                                self._region_labels_set[page_index]['items'][i] = True
                            print(f"[DEBUG] Added page number {page_index+1} to items DataFrame {i+1}")
                            items_df.append(df)

                    print(f"[DEBUG] Returning items data as a list with {len(items_df)} DataFrames for consistency")

            # Extract summary data
            if 'summary' in table_areas and table_areas['summary']:
                summary_tables = []
                for i, area in enumerate(table_areas['summary']):
                    # Get columns for this specific region
                    columns = None
                    if 'summary' in columns_list and i < len(columns_list['summary']):
                        columns = columns_list['summary'][i]
                        print(f"[DEBUG] Using columns for summary region {i}: {columns}")
                    # Get params without defaults
                    params = self.extraction_params.get('summary', {})

                    # Create a minimal set of parameters without defaults
                    extraction_params = {
                        'pages': str(page_index + 1),  # 1-based page number
                        'header': {'pages': str(page_index + 1)},
                        'items': {'pages': str(page_index + 1)},
                        'summary': {'pages': str(page_index + 1)}
                    }

                    # Determine the flavor - default to 'stream' if not specified
                    flavor = params.get('flavor', 'stream')

                    # Check if we have columns and need to adjust the flavor
                    if columns and flavor == 'lattice':
                        print(f"[WARNING] Columns are defined but flavor is set to 'lattice'. Changing to 'stream' flavor.")
                        flavor = 'stream'

                    # Set the flavor parameter
                    extraction_params['flavor'] = flavor

                    # Only add other parameters that exist in params
                    for key, value in params.items():
                        if key != 'flavor':  # Skip flavor as we've already handled it
                            extraction_params[key] = value

                    print(f"[DEBUG] extract_table called for summary on page {page_index + 1}")
                    print(f"[DEBUG] PDF path: {self.pdf_path}")
                    print(f"[DEBUG] Table area: {area}")
                    print(f"[DEBUG] Columns: {columns}")
                    print(f"[DEBUG] Using extraction params: {extraction_params}")

                    # Prepare additional parameters from custom parameters
                    additional_params = {}
                    for i in range(1, 4):  # Check for up to 3 custom parameters
                        param_name_key = f'custom_param_{i}_name'
                        param_value_key = f'custom_param_{i}_value'

                        if param_name_key in self.extraction_params and param_value_key in self.extraction_params:
                            param_name = self.extraction_params[param_name_key]
                            param_value = self.extraction_params[param_value_key]

                            if param_name and param_value is not None:
                                additional_params[param_name] = param_value
                                print(f"[DEBUG] Using custom parameter for summary: {param_name} = {param_value}")

                    # Print extraction parameters before calling extract_table
                    print(f"[DEBUG] Calling extract_table for summary with extraction_params: {self.extraction_params}")
                    print(f"[DEBUG] Summary section parameters: {self.extraction_params.get('summary', {})}")

                    # Extract table using multi-method extraction
                    extraction_method = getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
                    df = extract_with_method(
                        pdf_path=self.pdf_path,
                        extraction_method=extraction_method,
                        page_number=page_index + 1,  # 1-based page number
                        table_areas=[area],
                        columns_list=[columns] if columns else None,
                        section_type='summary',
                        extraction_params=self.extraction_params,  # Use the actual extraction_params from the class
                        use_cache=False  # Disable caching to ensure fresh extraction with column lines
                    )

                    if df is not None and not df.empty:
                        summary_tables.append(df)

                # Combine summary tables
                if summary_tables:
                    # Always return a list for consistency across all pages
                    # This ensures that page 1 and page 2 both have the same data structure
                    summary_df = []

                    # Process each table
                    for i, df in enumerate(summary_tables):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            # Add page number columns
                            df['page_number'] = page_index + 1  # 1-based page number
                            df['_page_number'] = page_index + 1  # Internal tracking

                            # Add region labels if they don't exist
                            if 'region_label' not in df.columns:
                                # Check if region labels have already been set for this page and section
                                if (page_index in self._region_labels_set and
                                    'summary' in self._region_labels_set[page_index] and
                                    i in self._region_labels_set[page_index]['summary']):
                                    print(f"[DEBUG] Region labels already set for summary DataFrame {i+1} on page {page_index+1}, skipping")
                                    continue

                                # Use the actual region label from the region dictionary
                                region_label = f"S{i+1}"  # Default
                                if 'summary' in self.page_regions.get(page_index, {}) and i < len(self.page_regions[page_index]['summary']):
                                    # Get the label from the corresponding summary region
                                    region_item = self.page_regions[page_index]['summary'][i]
                                    if isinstance(region_item, dict) and 'label' in region_item:
                                        region_label = region_item['label']

                                # Create region labels with page number to ensure uniqueness
                                # IMPORTANT: Preserve the original region label (S1, S2, etc.) exactly as it is
                                df['region_label'] = [f"{region_label}_R{j+1}_P{page_index+1}"
                                                    for j in range(len(df))]
                                print(f"[DEBUG] Created region labels for summary DataFrame {i+1}: {df['region_label'].tolist()}")
                                print(f"[DEBUG] Using original region label: {region_label} (preserving region number)")

                                # Mark these region labels as set
                                if page_index not in self._region_labels_set:
                                    self._region_labels_set[page_index] = {}
                                if 'summary' not in self._region_labels_set[page_index]:
                                    self._region_labels_set[page_index]['summary'] = {}
                                self._region_labels_set[page_index]['summary'][i] = True
                            else:
                                # DO NOT modify region labels - preserve them exactly as they are
                                # Just verify that the labels are present and log them
                                labels = df['region_label'].tolist()
                                print(f"[DEBUG] Preserving original region labels for summary DataFrame {i+1}: {labels}")
                                # No modification to the labels

                                # Mark these region labels as set
                                if page_index not in self._region_labels_set:
                                    self._region_labels_set[page_index] = {}
                                if 'summary' not in self._region_labels_set[page_index]:
                                    self._region_labels_set[page_index]['summary'] = {}
                                self._region_labels_set[page_index]['summary'][i] = True
                            print(f"[DEBUG] Added page number {page_index+1} to summary DataFrame {i+1}")
                            summary_df.append(df)

                    print(f"[DEBUG] Returning summary data as a list with {len(summary_df)} DataFrames for consistency")

            # Restore original page if needed
            if current_page != page_index:
                print(f"[DEBUG] Restoring to page {current_page + 1} after extraction")
                self.current_page_index = current_page
                self.display_current_page()

            return header_df, items_df, summary_df

        except Exception as e:
            print(f"Error extracting page data: {str(e)}")
            import traceback
            traceback.print_exc()

            # Restore original page if needed, even after an error
            if current_page != page_index:
                print(f"[DEBUG] Restoring to page {current_page + 1} after extraction error")
                self.current_page_index = current_page
                self.display_current_page()

            return None, None, None

    def _update_page_data(self):
        """Helper method to update the JSON tree with page data, handling None values safely"""
        print(f"[DEBUG] _update_page_data called for page {self.current_page_index + 1}")

        # Save current page's regions and column lines to ensure they're preserved
        if self.multi_page_mode:
            # Save current regions and column lines
            self.page_regions[self.current_page_index] = self.regions.copy()
            self.page_column_lines[self.current_page_index] = self.column_lines.copy()
            print(f"[DEBUG] Saved regions and column lines for page {self.current_page_index + 1} in _update_page_data")

        # In multi-page mode, use combined data from all pages unless we're in a specific region extraction
        if self.multi_page_mode and hasattr(self, 'pdf_document') and len(self.pdf_document) > 1:
            # Check if we're in a specific region extraction context
            specific_region_extraction = hasattr(self, '_in_specific_region_extraction') and self._in_specific_region_extraction

            # Also check if we should skip extraction update
            skip_extraction = hasattr(self, '_skip_extraction_update') and self._skip_extraction_update

            if skip_extraction:
                print(f"[DEBUG] Skipping extraction update in _update_page_data due to _skip_extraction_update flag")
                return True
            elif not specific_region_extraction:
                print(f"[DEBUG] Multi-page mode detected in _update_page_data, using combined data from all pages")

                # Always ensure _all_pages_data is properly initialized
                if not hasattr(self, '_all_pages_data'):
                    print(f"[DEBUG] _all_pages_data attribute not found, initializing in _update_page_data")
                    self._all_pages_data = [None] * len(self.pdf_document)
                elif not self._all_pages_data:
                    print(f"[DEBUG] _all_pages_data is empty, initializing in _update_page_data")
                    self._all_pages_data = [None] * len(self.pdf_document)
                elif len(self._all_pages_data) != len(self.pdf_document):
                    print(f"[DEBUG] _all_pages_data length mismatch, reinitializing in _update_page_data")
                    self._all_pages_data = [None] * len(self.pdf_document)

                # Make sure we have data for all pages before combining
                for page_idx in range(len(self.pdf_document)):
                    if page_idx >= len(self._all_pages_data) or self._all_pages_data[page_idx] is None:
                        print(f"[DEBUG] Extracting data for page {page_idx + 1} before combining in _update_page_data")
                        # Extract data for this page
                        header_df, items_df, summary_df = self.extract_page_data(page_idx)

                        # Store the extracted data in _all_pages_data
                        page_data = {
                            'header': header_df,
                            'items': items_df,
                            'summary': summary_df
                        }
                        self._all_pages_data[page_idx] = page_data
                        print(f"[DEBUG] Stored extraction data for page {page_idx + 1} in _update_page_data")

                        # Print summary of _all_pages_data after extraction
                        self._print_all_pages_data_summary("_update_page_data")

                # Force update with combined data from all pages
                print(f"[DEBUG] Ensuring ALL data from ALL pages is preserved without duplicate checking in _update_page_data")

                # Store the current page data in _all_pages_data before combining
                header_df, items_df, summary_df = self.extract_page_data(self.current_page_index)
                self._all_pages_data[self.current_page_index] = {
                    'header': header_df,
                    'items': items_df,
                    'summary': summary_df
                }

                # Now combine data from all pages
                # CRITICAL: Use cached multi-page extraction results if available
                # This ensures that all regions from all pages are preserved
                from pdf_extraction_utils import get_multipage_extraction
                cached_data = get_multipage_extraction(self.pdf_path)
                if cached_data:
                    print(f"[DEBUG] Using cached multi-page extraction results in _update_page_data")
                    combined_data = cached_data
                else:
                    # If no cached data is available, extract and combine data from all pages
                    combined_data = self.extract_multi_page_invoice()
                print(f"[DEBUG] Extracted combined data from all pages in _update_page_data")

                # Print summary of combined data to verify all regions are preserved
                for section in ['header', 'items', 'summary']:
                    if section in combined_data and isinstance(combined_data[section], pd.DataFrame) and not combined_data[section].empty:
                        if 'region_label' in combined_data[section].columns:
                            region_labels = combined_data[section]['region_label'].tolist()
                            print(f"[DEBUG] Combined {section} region labels in _update_page_data: {region_labels}")
                        if 'page_number' in combined_data[section].columns:
                            page_numbers = combined_data[section]['page_number'].unique().tolist()
                            print(f"[DEBUG] Combined {section} page numbers in _update_page_data: {page_numbers}")
                        print(f"[DEBUG] Combined {section} data shape in _update_page_data: {combined_data[section].shape}")
                        print(f"[DEBUG] Combined {section} row count in _update_page_data: {len(combined_data[section])}")
                    elif section in combined_data and isinstance(combined_data[section], list) and combined_data[section]:
                        print(f"[DEBUG] Combined {section} is a list with {len(combined_data[section])} items in _update_page_data")
                        for i, df in enumerate(combined_data[section]):
                            if isinstance(df, pd.DataFrame) and not df.empty:
                                if 'region_label' in df.columns:
                                    region_labels = df['region_label'].tolist()
                                    print(f"[DEBUG] Combined {section}[{i}] region labels in _update_page_data: {region_labels}")
                                if 'page_number' in df.columns:
                                    page_numbers = df['page_number'].unique().tolist()
                                    print(f"[DEBUG] Combined {section}[{i}] page numbers in _update_page_data: {page_numbers}")
                                print(f"[DEBUG] Combined {section}[{i}] data shape in _update_page_data: {df.shape}")
                                print(f"[DEBUG] Combined {section}[{i}] row count in _update_page_data: {len(df)}")
                    else:
                        print(f"[DEBUG] Combined {section} is empty or None in _update_page_data")

                # Add metadata
                if self.pdf_path:
                    combined_data['metadata'] = {
                        'filename': os.path.basename(self.pdf_path),
                        'page_count': len(self.pdf_document),
                        'template_type': 'multi',  # Always 'multi' for multi-page PDFs
                        'creation_date': datetime.datetime.now().isoformat()
                    }

                # Update the JSON tree with the combined data
                self.update_json_tree(combined_data)
                print(f"[DEBUG] Updated JSON tree with combined data from all pages in _update_page_data")
                return True
            else:
                print(f"[DEBUG] In specific region extraction context, skipping combined data extraction")
                # Continue with normal flow to use page-specific data

        # For single page mode or when not in multi-page mode
        if hasattr(self, '_all_pages_data') and self._all_pages_data is not None and self.current_page_index < len(self._all_pages_data):
            page_data = self._all_pages_data[self.current_page_index]
            if page_data is not None:
                # Make a deep copy to avoid modifying the original data
                import copy
                page_data_copy = copy.deepcopy(page_data)

                # DO NOT modify region labels - preserve them exactly as they are
                for section in ['header', 'items', 'summary']:
                    if section in page_data_copy:
                        if isinstance(page_data_copy[section], pd.DataFrame) and not page_data_copy[section].empty:
                            if 'region_label' in page_data_copy[section].columns:
                                # Just verify that the labels are present and log them
                                labels = page_data_copy[section]['region_label'].tolist()
                                print(f"[DEBUG] Preserving original region labels for {section} on page {self.current_page_index + 1}:")
                                print(f"[DEBUG] {labels}")
                                # No modification to the labels
                        elif isinstance(page_data_copy[section], list):
                            # Handle list of DataFrames
                            for i, df in enumerate(page_data_copy[section]):
                                if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                                    # Just verify that the labels are present and log them
                                    labels = df['region_label'].tolist()
                                    print(f"[DEBUG] Preserving original region labels for {section}[{i}] on page {self.current_page_index + 1}:")
                                    print(f"[DEBUG] {labels}")
                                    # No modification to the labels

                # Update the JSON tree with the modified data
                self.update_json_tree(page_data_copy)
                print(f"[DEBUG] Updated extraction results from stored data for page {self.current_page_index + 1}")
                return True
            else:
                print(f"[DEBUG] No stored data available for page {self.current_page_index + 1}, forcing extraction")
                # Force extraction for the current page if no stored data is available
                self.update_extraction_results(force=True)
                return True
        else:
            # Make sure _all_pages_data is initialized
            if not hasattr(self, '_all_pages_data') or not self._all_pages_data:
                self._all_pages_data = [None] * len(self.pdf_document)
                print(f"[DEBUG] Initialized _all_pages_data array with {len(self.pdf_document)} pages")

            # Force extraction for the current page if no stored data is available
            print(f"[DEBUG] No _all_pages_data available for page {self.current_page_index + 1}, forcing extraction")
            self.update_extraction_results(force=True)
            return True

    def show_tree_context_menu(self, position):
        """Show context menu for the extraction results view

        Args:
            position (QPoint): The position where the context menu should be shown
        """
        # Create a context menu
        menu = QMenu()

        # Get the current tab
        current_tab_index = self.extraction_tabs.currentIndex()
        current_tab = self.extraction_tabs.widget(current_tab_index)
        tab_name = self.extraction_tabs.tabText(current_tab_index)

        # Add copy action
        copy_action = menu.addAction("Copy")
        copy_action.triggered.connect(lambda: self.copy_selected_text(current_tab))

        # Add copy all action
        copy_all_action = menu.addAction("Copy All")
        copy_all_action.triggered.connect(lambda: self.copy_all_text(current_tab))

        # Add search action
        search_action = menu.addAction("Search...")
        search_action.triggered.connect(self.show_search_dialog)

        # Add separator
        menu.addSeparator()

        # Add export action
        export_action = menu.addAction("Export as JSON...")
        export_action.triggered.connect(self.export_as_json)

        # Show the menu at the given position
        menu.exec_(current_tab.mapToGlobal(position))

    def copy_selected_text(self, text_edit):
        """Copy selected text to clipboard

        Args:
            text_edit (QTextEdit): The text edit widget to copy from
        """
        selected_text = text_edit.textCursor().selectedText()
        if selected_text:
            clipboard = QApplication.clipboard()
            clipboard.setText(selected_text)
            print(f"[DEBUG] Copied selected text to clipboard ({len(selected_text)} characters)")

    def copy_all_text(self, text_edit):
        """Copy all text to clipboard

        Args:
            text_edit (QTextEdit): The text edit widget to copy from
        """
        all_text = text_edit.toPlainText()
        if all_text:
            clipboard = QApplication.clipboard()
            clipboard.setText(all_text)
            print(f"[DEBUG] Copied all text to clipboard ({len(all_text)} characters)")

    def export_as_json(self):
        """Export the extraction results as a JSON file"""
        # Get the current data
        current_data = self._get_current_json_data()

        if not current_data:
            QMessageBox.warning(self, "Export Error", "No data available to export")
            return

        # Ask for a file name
        file_dialog = QFileDialog()
        file_dialog.setAcceptMode(QFileDialog.AcceptSave)
        file_dialog.setNameFilter("JSON Files (*.json)")
        file_dialog.setDefaultSuffix("json")

        # Set a default file name based on the PDF file name if available
        default_name = "extraction_results.json"
        if hasattr(self, 'pdf_path') and self.pdf_path:
            pdf_name = os.path.basename(self.pdf_path)
            default_name = os.path.splitext(pdf_name)[0] + "_results.json"

        file_dialog.selectFile(default_name)

        if file_dialog.exec_() == QDialog.Accepted:
            file_path = file_dialog.selectedFiles()[0]

            try:
                # Convert DataFrames to lists of dictionaries for JSON serialization
                json_data = {}

                # Add metadata
                if 'metadata' in current_data:
                    json_data['metadata'] = current_data['metadata']

                # Process header data
                if 'header' in current_data and current_data['header'] is not None:
                    if isinstance(current_data['header'], list):
                        json_data['header'] = []
                        for df in current_data['header']:
                            if isinstance(df, pd.DataFrame) and not df.empty:
                                json_data['header'].append(df.to_dict('records'))
                    elif isinstance(current_data['header'], pd.DataFrame) and not current_data['header'].empty:
                        json_data['header'] = current_data['header'].to_dict('records')
                    else:
                        json_data['header'] = []
                else:
                    json_data['header'] = []

                # Process items data
                if 'items' in current_data and current_data['items'] is not None:
                    if isinstance(current_data['items'], list):
                        json_data['items'] = []
                        for df in current_data['items']:
                            if isinstance(df, pd.DataFrame) and not df.empty:
                                json_data['items'].append(df.to_dict('records'))
                    elif isinstance(current_data['items'], pd.DataFrame) and not current_data['items'].empty:
                        json_data['items'] = current_data['items'].to_dict('records')
                    else:
                        json_data['items'] = []
                else:
                    json_data['items'] = []

                # Process summary data
                if 'summary' in current_data and current_data['summary'] is not None:
                    if isinstance(current_data['summary'], list):
                        json_data['summary'] = []
                        for df in current_data['summary']:
                            if isinstance(df, pd.DataFrame) and not df.empty:
                                json_data['summary'].append(df.to_dict('records'))
                    elif isinstance(current_data['summary'], pd.DataFrame) and not current_data['summary'].empty:
                        json_data['summary'] = current_data['summary'].to_dict('records')
                    else:
                        json_data['summary'] = []
                else:
                    json_data['summary'] = []

                # Write to file
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, indent=2, ensure_ascii=False)

                print(f"[DEBUG] Exported extraction results to {file_path}")

                # Show success message
                QMessageBox.information(
                    self,
                    "Export Successful",
                    f"Extraction results exported to:\n{file_path}"
                )

                # Open the file in the default application
                try:
                    os.startfile(file_path)
                except Exception as e:
                    print(f"[DEBUG] Error opening exported file: {str(e)}")

            except Exception as e:
                print(f"[ERROR] Failed to export as JSON: {str(e)}")
                import traceback
                traceback.print_exc()

                QMessageBox.warning(
                    self,
                    "Export Error",
                    f"Failed to export as JSON: {str(e)}"
                )

    def show_search_dialog(self):
        """Show a dialog to search for text in the extraction results"""
        # Get the current tab
        current_tab_index = self.extraction_tabs.currentIndex()
        current_tab = self.extraction_tabs.widget(current_tab_index)

        # Create a dialog for entering search text
        search_text, ok = QInputDialog.getText(
            self,
            "Search",
            "Enter text or regex pattern to search for:",
            QLineEdit.Normal,
            self._active_regex_pattern if hasattr(self, '_active_regex_pattern') else ""
        )

        if ok and search_text:
            # Highlight matches
            self._active_regex_pattern = search_text
            self.highlight_regex_matches(search_text)

    def highlight_regex_matches(self, regex_pattern):
        """Highlight text matching a regex pattern in the extraction results view

        Args:
            regex_pattern (str): The regex pattern to highlight
        """
        if not hasattr(self, 'extraction_tabs'):
            return

        # Get the current active tab
        current_tab_index = self.extraction_tabs.currentIndex()
        current_tab = self.extraction_tabs.widget(current_tab_index)

        # Clear previous highlighting in the current tab
        cursor = current_tab.textCursor()
        cursor.select(QTextCursor.Document)
        format = QTextCharFormat()
        cursor.setCharFormat(format)

        # If pattern is empty, just return after clearing
        if not regex_pattern or regex_pattern.strip() == "":
            return

        # Store the current pattern for future reference
        self._active_regex_pattern = regex_pattern

        try:
            # Create a regular expression object
            regex = QRegularExpression(regex_pattern)
            if not regex.isValid():
                print(f"[DEBUG] Invalid regex pattern: {regex_pattern}")
                return

            # Get the text content of the current tab
            text_content = current_tab.toPlainText()

            # Create a format for highlighting matches
            highlight_format = QTextCharFormat()
            highlight_format.setBackground(QBrush(QColor(0, 200, 0, 100)))  # Light green with transparency

            # Find all matches
            match_count = 0
            match_iterator = regex.globalMatch(text_content)
            while match_iterator.hasNext():
                match = match_iterator.next()
                if match.hasMatch():
                    # Create a cursor at the match position
                    cursor = QTextCursor(current_tab.document())
                    cursor.setPosition(match.capturedStart())
                    cursor.setPosition(match.capturedEnd(), QTextCursor.KeepAnchor)

                    # Apply the highlighting format
                    cursor.setCharFormat(highlight_format)
                    match_count += 1

            if match_count > 0:
                tab_name = self.extraction_tabs.tabText(current_tab_index)
                print(f"[DEBUG] Highlighted {match_count} matches for pattern: {regex_pattern} in {tab_name} tab")
        except Exception as e:
            print(f"[DEBUG] Error highlighting regex matches: {str(e)}")

    def update_json_tree(self, data):
        """Update the extraction results view with raw invoice2data format text"""
        if not hasattr(self, 'json_tree'):
            return

        # Check if data is None before trying to copy it
        if data is None:
            print(f"[WARNING] Received None data in update_json_tree, using empty structure")
            # If we have cached data, use it instead of creating an empty structure
            if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data is not None:
                print(f"[DEBUG] Using existing cached data instead of empty structure")
                data = self._cached_extraction_data.copy()
            else:
                # Initialize with consistent empty structure
                data = {
                    'header': [],  # Always use list for consistency
                    'items': [],   # Always use list for consistency
                    'summary': []  # Always use list for consistency
                }

        # Check if we're in a specific region extraction context
        # If we are, use the provided data directly without combining
        specific_region_extraction = hasattr(self, '_in_specific_region_extraction') and self._in_specific_region_extraction

        # Check if we should skip extraction update
        skip_extraction = hasattr(self, '_skip_extraction_update') and self._skip_extraction_update

        # If we should skip extraction update, just use the provided data directly
        if skip_extraction:
            print(f"[DEBUG] Skipping automatic extraction update due to _skip_extraction_update flag")
            # Make sure we're using the cached data if available
            if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data is not None:
                data = self._cached_extraction_data.copy()
                print(f"[DEBUG] Using cached extraction data: {list(data.keys())}")
        # If we're in a specific region extraction context, use the provided data directly
        elif specific_region_extraction:
            print(f"[DEBUG] In specific region extraction context, using provided data directly")
            # Update the cached data with the new data
            if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data is not None:
                # Merge the new data with the cached data
                for section in data.keys():
                    if section in self._cached_extraction_data:
                        self._cached_extraction_data[section] = data[section]
                data = self._cached_extraction_data.copy()
                print(f"[DEBUG] Updated cached data with specific region data: {list(data.keys())}")
        # Use combined data from all pages in multi-page mode
        elif self.multi_page_mode and hasattr(self, 'pdf_document') and len(self.pdf_document) > 1:
            print(f"[DEBUG] Multi-page mode detected in update_json_tree, using combined data from all pages")
            print(f"[DEBUG] Current page index: {self.current_page_index + 1}")
            print(f"[DEBUG] PDF has {len(self.pdf_document)} pages")

            # Always ensure _all_pages_data is properly initialized
            if not hasattr(self, '_all_pages_data'):
                print(f"[DEBUG] _all_pages_data attribute not found, initializing")
                self._all_pages_data = [None] * len(self.pdf_document)
            elif not self._all_pages_data:
                print(f"[DEBUG] _all_pages_data is empty, initializing")
                self._all_pages_data = [None] * len(self.pdf_document)
            elif len(self._all_pages_data) != len(self.pdf_document):
                print(f"[DEBUG] _all_pages_data length mismatch, reinitializing")
                self._all_pages_data = [None] * len(self.pdf_document)

            # Check if we need to extract data for any pages
            missing_pages = []
            for page_idx in range(len(self.pdf_document)):
                if page_idx >= len(self._all_pages_data) or self._all_pages_data[page_idx] is None:
                    missing_pages.append(page_idx)

            if missing_pages:
                print(f"[DEBUG] Need to extract data for pages: {[p+1 for p in missing_pages]}")
                for page_idx in missing_pages:
                    print(f"[DEBUG] Extracting data for page {page_idx + 1}")
                    # Extract data for this page
                    header_df, items_df, summary_df = self.extract_page_data(page_idx)

                    # Store the extracted data in _all_pages_data
                    page_data = {
                        'header': header_df,
                        'items': items_df,
                        'summary': summary_df
                    }
                    self._all_pages_data[page_idx] = page_data
                    print(f"[DEBUG] Stored extraction data for page {page_idx + 1}")

            # Store the current page data in _all_pages_data before combining
            if data is not None and isinstance(data, dict):
                # Always update the current page data to ensure it's preserved
                self._all_pages_data[self.current_page_index] = data.copy()
                print(f"[DEBUG] Updated _all_pages_data for page {self.current_page_index + 1} with new data")

            # Now combine data from all pages
            print(f"[DEBUG] Using _all_pages_data to combine data from all pages")
            print(f"[DEBUG] Ensuring ALL data from ALL pages is preserved without duplicate checking")
            combined_data = self.extract_multi_page_invoice()
            print(f"[DEBUG] Combined data from all pages using extract_multi_page_invoice")

            # Add metadata only if it doesn't already exist
            if self.pdf_path:
                if 'metadata' not in combined_data:
                    combined_data['metadata'] = {
                        'filename': os.path.basename(self.pdf_path),
                        'page_count': len(self.pdf_document),
                        'template_type': 'multi',  # Always 'multi' for multi-page PDFs
                        'creation_date': datetime.datetime.now().isoformat()
                    }
                    print(f"[DEBUG] Added metadata to combined data in update_json_tree with creation_date: {combined_data['metadata']['creation_date']}")
                else:
                    # Preserve the existing creation_date to prevent region label changes
                    print(f"[DEBUG] Metadata already exists in combined data in update_json_tree, preserving original creation_date: {combined_data['metadata'].get('creation_date', 'N/A')}")

            # Use the combined data instead of the page-specific data in multi-page mode
            data = combined_data
            print(f"[DEBUG] Using combined data from all pages for display")

            # CRITICAL: Ensure the original region labels are preserved exactly as they are
            # This is the key fix to prevent region labels from being changed
            for section in ['header', 'items', 'summary']:
                if section in data:
                    if isinstance(data[section], list):
                        for i, df in enumerate(data[section]):
                            if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                                # Verify that the labels are present and log them
                                labels = df['region_label'].tolist()
                                print(f"[DEBUG] Preserving original region labels for {section}[{i}]: {labels}")

                                # Make a deep copy of the original labels to ensure they're not modified
                                original_labels = df['region_label'].copy()

                                # Ensure the region labels are preserved exactly as they are
                                # This is critical to prevent region labels from changing
                                df['region_label'] = original_labels

                                # Log the preserved labels to verify they're unchanged
                                preserved_labels = df['region_label'].tolist()
                                print(f"[DEBUG] Verified preserved region labels for {section}[{i}]: {preserved_labels}")

                                # CRITICAL: Ensure page numbers are preserved exactly as they are
                                if 'page_number' in df.columns:
                                    # Make a deep copy of the original page numbers
                                    original_page_numbers = df['page_number'].copy()
                                    # Ensure the page numbers are preserved exactly as they are
                                    df['page_number'] = original_page_numbers
                                    print(f"[DEBUG] Preserved page numbers for {section}[{i}]: {df['page_number'].unique().tolist()}")
                    elif isinstance(data[section], pd.DataFrame) and not data[section].empty and 'region_label' in data[section].columns:
                        # Verify that the labels are present and log them
                        labels = data[section]['region_label'].tolist()
                        print(f"[DEBUG] Preserving original region labels for {section}: {labels}")

                        # Make a deep copy of the original labels to ensure they're not modified
                        original_labels = data[section]['region_label'].copy()

                        # Ensure the region labels are preserved exactly as they are
                        # This is critical to prevent region labels from changing
                        data[section]['region_label'] = original_labels

                        # Log the preserved labels to verify they're unchanged
                        preserved_labels = data[section]['region_label'].tolist()
                        print(f"[DEBUG] Verified preserved region labels for {section}: {preserved_labels}")

                        # CRITICAL: Ensure page numbers are preserved exactly as they are
                        if 'page_number' in data[section].columns:
                            # Make a deep copy of the original page numbers
                            original_page_numbers = data[section]['page_number'].copy()
                            # Ensure the page numbers are preserved exactly as they are
                            data[section]['page_number'] = original_page_numbers
                            print(f"[DEBUG] Preserved page numbers for {section}: {data[section]['page_number'].unique().tolist()}")

            # Ensure header data is preserved as a list of DataFrames by page
            if 'header' in combined_data and isinstance(combined_data['header'], pd.DataFrame) and not combined_data['header'].empty:
                # Group header data by page_number
                if 'page_number' in combined_data['header'].columns:
                    header_by_page = []
                    for page_num, group in combined_data['header'].groupby('page_number'):
                        # Preserve the original region labels exactly as they are
                        if 'region_label' in group.columns:
                            print(f"[DEBUG] Preserving original region labels for header on page {page_num}")
                            print(f"[DEBUG] Original region labels: {group['region_label'].tolist()}")

                        header_by_page.append(group.reset_index(drop=True))
                        print(f"[DEBUG] Created header group for page {page_num} with {len(group)} rows")

                    # Replace the single DataFrame with a list of DataFrames
                    combined_data['header'] = header_by_page
                    data['header'] = header_by_page
                    print(f"[DEBUG] Converted header from single DataFrame to list of {len(header_by_page)} DataFrames by page")

            # Print summary of combined data to verify all regions are preserved
            for section in ['header', 'items', 'summary']:
                if section in combined_data and isinstance(combined_data[section], pd.DataFrame) and not combined_data[section].empty:
                    if 'region_label' in combined_data[section].columns:
                        region_labels = combined_data[section]['region_label'].tolist()
                        print(f"[DEBUG] Combined {section} region labels: {region_labels}")
                    if 'page_number' in combined_data[section].columns:
                        page_numbers = combined_data[section]['page_number'].unique().tolist()
                        print(f"[DEBUG] Combined {section} page numbers: {page_numbers}")
                    print(f"[DEBUG] Combined {section} data shape: {combined_data[section].shape}")
                    print(f"[DEBUG] Combined {section} row count: {len(combined_data[section])}")
                elif section in combined_data and isinstance(combined_data[section], list) and combined_data[section]:
                    print(f"[DEBUG] Combined {section} is a list with {len(combined_data[section])} items")
                    for i, df in enumerate(combined_data[section]):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            if 'region_label' in df.columns:
                                region_labels = df['region_label'].tolist()
                                print(f"[DEBUG] Combined {section}[{i}] region labels: {region_labels}")
                            if 'page_number' in df.columns:
                                page_numbers = df['page_number'].unique().tolist()
                                print(f"[DEBUG] Combined {section}[{i}] page numbers: {page_numbers}")
                            print(f"[DEBUG] Combined {section}[{i}] data shape: {df.shape}")
                            print(f"[DEBUG] Combined {section}[{i}] row count: {len(df)}")
                else:
                    print(f"[DEBUG] Combined {section} is empty or None")

        # Cache the data for future updates - only if it's not empty
        if data is not None and isinstance(data, dict) and any(
            section is not None and (
                (isinstance(section, pd.DataFrame) and not section.empty) or
                (not isinstance(section, (list, pd.DataFrame)) and section) or
                (isinstance(section, list) and len(section) > 0)
            ) for section in data.values()
        ):
            self._cached_extraction_data = data.copy()
            print(f"[DEBUG] Cached extraction data: {list(data.keys())}")
        else:
            print(f"[WARNING] Not caching empty data in update_json_tree")

        # Add page information to the data if it's not already there
        if 'page_info' not in data and self.multi_page_mode:
            # Check if any of the DataFrames have page_number columns
            has_page_info = False
            for section in ['header', 'items', 'summary']:
                if section in data and isinstance(data[section], pd.DataFrame) and 'page_number' in data[section].columns:
                    has_page_info = True
                    break

            if has_page_info:
                # Create page_info structure
                data['page_info'] = {}
                print(f"[DEBUG] Adding page_info to extraction data")

                # Add page info for header
                if 'header' in data and isinstance(data['header'], pd.DataFrame) and 'page_number' in data['header'].columns:
                    page_nums = data['header']['page_number'].unique()
                    if len(page_nums) > 0:
                        data['page_info']['header'] = {'page_number': int(page_nums[0])}
                        print(f"[DEBUG] Header is from page {int(page_nums[0])}")

                # Add page info for items
                if 'items' in data and isinstance(data['items'], pd.DataFrame) and 'page_number' in data['items'].columns:
                    page_nums = data['items']['page_number'].unique()
                    for i, page_num in enumerate(page_nums):
                        data['page_info'][f'items_{i}'] = {'page_number': int(page_num)}
                        print(f"[DEBUG] Items section {i+1} is from page {int(page_num)}")

                # Add page info for summary
                if 'summary' in data and isinstance(data['summary'], pd.DataFrame) and 'page_number' in data['summary'].columns:
                    page_nums = data['summary']['page_number'].unique()
                    for i, page_num in enumerate(page_nums):
                        data['page_info'][f'summary_{i}'] = {'page_number': int(page_num)}
                        print(f"[DEBUG] Summary section {i+1} is from page {int(page_num)}")

        # Ensure all DataFrames have page_number columns
        for section in ['header', 'items', 'summary']:
            if section in data:
                if isinstance(data[section], pd.DataFrame) and not data[section].empty:
                    if 'page_number' not in data[section].columns:
                        print(f"[DEBUG] Adding missing page_number column to {section} DataFrame")
                        data[section]['page_number'] = self.current_page_index + 1  # Default to current page
                    if '_page_number' not in data[section].columns:
                        print(f"[DEBUG] Adding missing _page_number column to {section} DataFrame")
                        data[section]['_page_number'] = self.current_page_index + 1  # Default to current page
                elif isinstance(data[section], list):
                    # Handle list of DataFrames
                    for i, df in enumerate(data[section]):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            if 'page_number' not in df.columns:
                                print(f"[DEBUG] Adding missing page_number column to {section}[{i}] DataFrame")
                                df['page_number'] = self.current_page_index + 1  # Default to current page
                            if '_page_number' not in df.columns:
                                print(f"[DEBUG] Adding missing _page_number column to {section}[{i}] DataFrame")
                                df['_page_number'] = self.current_page_index + 1  # Default to current page

        # Convert the data to invoice2data format text using the unified invoice processing utilities
        raw_text = invoice_processing_utils.convert_extraction_to_text(data, pdf_path=self.pdf_path)

        # Create a tabbed display widget if it doesn't exist
        if not hasattr(self, 'extraction_tabs') or self.extraction_tabs is None:
            # Remove the tree widget from its parent layout
            parent_widget = self.json_tree.parent()
            parent_layout = parent_widget.layout()

            # Create a splitter to divide the extraction results section horizontally
            self.extraction_splitter = QSplitter(Qt.Vertical, parent_widget)

            # Create top and bottom tab widgets
            self.top_tabs = QTabWidget()
            self.top_tabs.currentChanged.connect(self.on_top_tab_changed)
            self.bottom_tabs = QTabWidget()
            self.bottom_tabs.currentChanged.connect(self.on_bottom_tab_changed)

            # Create tabs for each section in the top half (raw text)
            # Metadata tab
            self.metadata_tab = QTextEdit()
            self.metadata_tab.setReadOnly(True)
            self.metadata_tab.setFont(QFont("Courier New", 10))
            self.metadata_tab.setLineWrapMode(QTextEdit.NoWrap)

            # Create header raw text tab for top section
            self.header_raw_text_tab = QTextEdit()
            self.header_raw_text_tab.setReadOnly(True)
            self.header_raw_text_tab.setFont(QFont("Courier New", 10))
            self.header_raw_text_tab.setLineWrapMode(QTextEdit.NoWrap)

            # Create items raw text tab for top section
            self.items_raw_text_tab = QTextEdit()
            self.items_raw_text_tab.setReadOnly(True)
            self.items_raw_text_tab.setFont(QFont("Courier New", 10))
            self.items_raw_text_tab.setLineWrapMode(QTextEdit.NoWrap)

            # Create summary raw text tab for top section
            self.summary_raw_text_tab = QTextEdit()
            self.summary_raw_text_tab.setReadOnly(True)
            self.summary_raw_text_tab.setFont(QFont("Courier New", 10))
            self.summary_raw_text_tab.setLineWrapMode(QTextEdit.NoWrap)

            # Add tabs to the top tab widget
            self.top_tabs.addTab(self.metadata_tab, "Metadata")
            self.top_tabs.addTab(self.header_raw_text_tab, "Header")
            self.top_tabs.addTab(self.items_raw_text_tab, "Items")
            self.top_tabs.addTab(self.summary_raw_text_tab, "Summary")

            # Create a common tab widget for the bottom section
            self.common_tab_widget = QTabWidget()

            # Create Fields tab (common for all sections)
            self.fields_editor = QWidget()
            fields_layout = QVBoxLayout(self.fields_editor)

            # Add a note about required fields
            required_fields_note = QLabel("Required fields: invoice_number, date, amount")
            required_fields_note.setStyleSheet("font-weight: bold; color: #d32f2f;")
            required_fields_note.setAlignment(Qt.AlignCenter)
            fields_layout.addWidget(required_fields_note)

            # Add a description of the required fields
            fields_description = QLabel(
                "Every template must have these fields:\n"
                "• invoice_number: Unique number assigned to invoice by issuer\n"
                "• date: Invoice issue date\n"
                "• amount: Total amount (with taxes)"
            )
            fields_description.setWordWrap(True)
            fields_description.setStyleSheet("font-style: italic; margin-bottom: 10px;")
            fields_layout.addWidget(fields_description)

            # Table for fields
            self.header_fields_table = QTableWidget(0, 3)
            self.header_fields_table.setHorizontalHeaderLabels(["Field Name", "Regex Pattern", "Type"])
            self.header_fields_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            fields_layout.addWidget(self.header_fields_table)

            # Create a reference to fields_table for backward compatibility
            self.fields_table = self.header_fields_table

            # Buttons for fields
            fields_buttons = QHBoxLayout()
            add_field_btn = QPushButton("Add Field")
            add_field_btn.clicked.connect(self.add_field)
            remove_field_btn = QPushButton("Remove Selected")
            remove_field_btn.clicked.connect(self.remove_field)
            fields_buttons.addWidget(add_field_btn)
            fields_buttons.addWidget(remove_field_btn)
            fields_layout.addLayout(fields_buttons)

            # Create Tables tab (common for all sections)
            self.tables_editor = QWidget()
            tables_layout = QVBoxLayout(self.tables_editor)

            # Add a description of the tables plugin
            tables_description = QLabel(
                "The tables plugin allows you to extract data from tables where the column headers \n"
                "and their corresponding values are on different lines. This is often the case in \n"
                "invoices where data is presented in a more visual, tabular format."
            )
            tables_description.setWordWrap(True)
            tables_description.setStyleSheet("font-style: italic; margin-bottom: 10px;")
            tables_layout.addWidget(tables_description)

            # Table for tables definitions
            self.tables_definition_table = QTableWidget(0, 3)
            self.tables_definition_table.setHorizontalHeaderLabels(["Start Pattern", "End Pattern", "Body Pattern"])
            self.tables_definition_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            self.tables_definition_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            self.tables_definition_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            tables_layout.addWidget(self.tables_definition_table)

            # Buttons for tables
            tables_buttons = QHBoxLayout()
            add_table_btn = QPushButton("Add Table Definition")
            add_table_btn.clicked.connect(self.add_table_definition)
            remove_table_btn = QPushButton("Remove Selected")
            remove_table_btn.clicked.connect(self.remove_table_definition)
            tables_buttons.addWidget(add_table_btn)
            tables_buttons.addWidget(remove_table_btn)
            tables_layout.addLayout(tables_buttons)

            # Create Lines tab (common for all sections)
            self.lines_editor = QWidget()
            lines_layout = QFormLayout(self.lines_editor)

            self.line_start_input = QLineEdit()
            self.line_start_input.textChanged.connect(self.on_regex_pattern_changed)
            lines_layout.addRow("Start Pattern:", self.line_start_input)

            self.line_end_input = QLineEdit()
            self.line_end_input.textChanged.connect(self.on_regex_pattern_changed)
            lines_layout.addRow("End Pattern:", self.line_end_input)

            self.line_pattern_input = QLineEdit()
            self.line_pattern_input.textChanged.connect(self.on_regex_pattern_changed)
            lines_layout.addRow("Line Pattern:", self.line_pattern_input)

            self.skip_line_input = QLineEdit()
            self.skip_line_input.textChanged.connect(self.on_regex_pattern_changed)
            lines_layout.addRow("Skip Line Pattern:", self.skip_line_input)

            # First line patterns (list widget with add/remove buttons)
            firstline_container = QWidget()
            firstline_layout = QVBoxLayout(firstline_container)
            firstline_layout.setContentsMargins(0, 0, 0, 0)

            self.firstline_list = QListWidget()
            firstline_layout.addWidget(self.firstline_list)

            firstline_buttons = QHBoxLayout()
            add_firstline_btn = QPushButton("Add Pattern")
            add_firstline_btn.clicked.connect(self.add_firstline_pattern)
            remove_firstline_btn = QPushButton("Remove Selected")
            remove_firstline_btn.clicked.connect(self.remove_firstline_pattern)
            firstline_buttons.addWidget(add_firstline_btn)
            firstline_buttons.addWidget(remove_firstline_btn)
            firstline_layout.addLayout(firstline_buttons)

            lines_layout.addRow("First Line Patterns:", firstline_container)

            # Line types table
            self.line_types_table = QTableWidget(0, 2)
            self.line_types_table.setHorizontalHeaderLabels(["Field Name", "Type"])
            lines_layout.addRow("Line Types:", self.line_types_table)

            # Buttons for line types
            linetypes_buttons = QHBoxLayout()
            add_linetype_btn = QPushButton("Add Line Type")
            add_linetype_btn.clicked.connect(self.add_line_type)
            remove_linetype_btn = QPushButton("Remove Selected")
            remove_linetype_btn.clicked.connect(self.remove_line_type)
            linetypes_buttons.addWidget(add_linetype_btn)
            linetypes_buttons.addWidget(remove_linetype_btn)

            # Add a container for the buttons
            linetypes_container = QWidget()
            linetypes_container.setLayout(linetypes_buttons)
            lines_layout.addRow("", linetypes_container)

            # Create Tax Lines tab (common for all sections)
            self.tax_lines_editor = QWidget()
            tax_lines_layout = QVBoxLayout(self.tax_lines_editor)

            # Add a description of the tax_lines section
            tax_lines_description = QLabel(
                "The tax_lines section allows you to extract tax information from invoices. \n"
                "This is often presented as a table near the bottom with a summary of the applied VAT taxes."
            )
            tax_lines_description.setWordWrap(True)
            tax_lines_description.setStyleSheet("font-style: italic; margin-bottom: 10px;")
            tax_lines_layout.addWidget(tax_lines_description)

            # Start and end patterns for tax lines
            self.tax_lines_start_input = QLineEdit()
            self.tax_lines_start_input.textChanged.connect(self.on_regex_pattern_changed)
            tax_lines_layout.addWidget(QLabel("Start Pattern:"))
            tax_lines_layout.addWidget(self.tax_lines_start_input)

            self.tax_lines_end_input = QLineEdit()
            self.tax_lines_end_input.textChanged.connect(self.on_regex_pattern_changed)
            tax_lines_layout.addWidget(QLabel("End Pattern:"))
            tax_lines_layout.addWidget(self.tax_lines_end_input)

            self.tax_lines_line_input = QLineEdit()
            self.tax_lines_line_input.textChanged.connect(self.on_regex_pattern_changed)
            tax_lines_layout.addWidget(QLabel("Line Pattern (with named capture groups):"))
            tax_lines_layout.addWidget(self.tax_lines_line_input)

            # Tax line types table
            self.tax_line_types_table = QTableWidget(0, 2)
            self.tax_line_types_table.setHorizontalHeaderLabels(["Field Name", "Type"])
            tax_lines_layout.addWidget(QLabel("Tax Line Types:"))
            tax_lines_layout.addWidget(self.tax_line_types_table)

            # Buttons for tax line types
            tax_line_types_buttons = QHBoxLayout()
            add_tax_line_type_btn = QPushButton("Add Tax Line Type")
            add_tax_line_type_btn.clicked.connect(self.add_tax_line_type)
            remove_tax_line_type_btn = QPushButton("Remove Selected")
            remove_tax_line_type_btn.clicked.connect(self.remove_tax_line_type)
            tax_line_types_buttons.addWidget(add_tax_line_type_btn)
            tax_line_types_buttons.addWidget(remove_tax_line_type_btn)
            tax_lines_layout.addLayout(tax_line_types_buttons)

            # Add tabs to the common tab widget
            self.common_tab_widget.addTab(self.fields_editor, "Fields")
            self.common_tab_widget.addTab(self.tables_editor, "Tables")
            self.common_tab_widget.addTab(self.lines_editor, "Lines")
            self.common_tab_widget.addTab(self.tax_lines_editor, "Tax Lines")

            # Create references to the old tab widgets for backward compatibility
            self.header_tab_widget = self.common_tab_widget
            self.items_tab_widget = self.common_tab_widget
            self.summary_tab_widget = self.common_tab_widget

            # Create references to the old editors for backward compatibility
            self.header_fields_editor = self.fields_editor
            self.items_tables_editor = self.tables_editor
            self.items_lines_editor = self.lines_editor
            self.summary_fields_editor = self.fields_editor

            # Add the common tab widget to the bottom tabs
            self.bottom_tabs.addTab(self.common_tab_widget, "JSON Designer")

            # Add context menu for copy functionality to all tabs
            for tab in [self.metadata_tab, self.header_raw_text_tab, self.items_raw_text_tab, self.summary_raw_text_tab]:
                tab.setContextMenuPolicy(Qt.CustomContextMenu)
                tab.customContextMenuRequested.connect(self.show_tree_context_menu)

            # Add the top and bottom tabs to the splitter
            self.extraction_splitter.addWidget(self.top_tabs)
            self.extraction_splitter.addWidget(self.bottom_tabs)

            # Set initial sizes for the splitter (50% each)
            self.extraction_splitter.setSizes([500, 500])

            # Replace the tree widget with the splitter
            parent_layout.replaceWidget(self.json_tree, self.extraction_splitter)

            # Hide the tree widget
            self.json_tree.hide()

            # Create a reference to the extraction tabs for backward compatibility
            self.extraction_tabs = self.top_tabs

            # Set the metadata tab as active by default
            self.top_tabs.setCurrentIndex(0)  # Metadata tab is at index 0
            print(f"[DEBUG] Set metadata tab as active by default")

        # Parse the raw text and distribute to respective tabs
        # The raw text is in the format:
        # METADATA
        # key1|value1
        # key2|value2
        # ...
        #
        # HEADER
        # header data...
        #
        # ITEMS
        # items data...
        #
        # SUMMARY
        # summary data...

        # Split the raw text into sections
        sections = {}
        current_section = None
        section_text = ""

        for line in raw_text.splitlines():
            if line.strip() in ["METADATA", "HEADER", "ITEMS", "SUMMARY"]:
                # Save the previous section if any
                if current_section:
                    sections[current_section] = section_text.strip()
                    section_text = ""

                # Start a new section
                current_section = line.strip()
            else:
                # Add the line to the current section
                section_text += line + "\n"

        # Save the last section
        if current_section:
            sections[current_section] = section_text.strip()

        # Update each section tab with only raw text data
        # Metadata tab - only raw text
        metadata_text = ""
        if "METADATA" in sections:
            metadata_text = sections["METADATA"]

        self.metadata_tab.setText(metadata_text)

        # Header tabs - raw text
        header_text = ""
        if "HEADER" in sections:
            header_text = sections["HEADER"]

        self.header_raw_text_tab.setText(header_text)  # Top tab

        # Items tabs - raw text
        items_text = ""
        if "ITEMS" in sections:
            items_text = sections["ITEMS"]

        self.items_raw_text_tab.setText(items_text)  # Top tab

        # Summary tabs - raw text
        summary_text = ""
        if "SUMMARY" in sections:
            summary_text = sections["SUMMARY"]

        self.summary_raw_text_tab.setText(summary_text)  # Top tab

        # Log the update
        print(f"[DEBUG] Updated extraction results with tabbed view ({len(raw_text.splitlines())} lines)")

        # Apply any active regex highlighting
        if hasattr(self, '_active_regex_pattern') and self._active_regex_pattern:
            self.highlight_regex_matches(self._active_regex_pattern)



    def update_json_tree_for_tables(self, table_list, parent_item=None):
        """Update the extraction results view with multiple tables in raw invoice2data format"""
        # Create a dictionary structure to hold the tables
        data = {
            'items': table_list  # Treat all tables as items for simplicity
        }

        # Use the update_json_tree method to display the raw text
        self.update_json_tree(data)

    def copy_all_data_to_clipboard(self):
        """Copy all data from the extraction results to clipboard in the same format as sent to invoice2data

        For multi-page PDFs, this will combine data from all pages by section before copying.
        Uses _all_pages_data directly to avoid redundant extraction.
        """
        # Check if we're in multi-page mode and need to combine data from all pages
        if self.multi_page_mode and hasattr(self, 'pdf_document') and len(self.pdf_document) > 1:
            print(f"[DEBUG] Multi-page mode detected, combining data from all pages for clipboard")

            # Use extract_multi_page_invoice which now uses the cache mechanism
            combined_data = self.extract_multi_page_invoice()

            # Add metadata only if it doesn't already exist
            if self.pdf_path:
                if 'metadata' not in combined_data:
                    combined_data['metadata'] = {
                        'filename': os.path.basename(self.pdf_path),
                        'page_count': len(self.pdf_document),
                        'template_type': 'multi',  # Always 'multi' for multi-page PDFs
                        'creation_date': datetime.datetime.now().isoformat()
                    }
                    print(f"[DEBUG] Added metadata to combined data in copy_all_data_to_clipboard")
                else:
                    print(f"[DEBUG] Metadata already exists in combined data in copy_all_data_to_clipboard, not updating creation_date")

                # CRITICAL: Ensure the original region labels are preserved exactly as they are
                for section in ['header', 'items', 'summary']:
                    if section in combined_data:
                        if isinstance(combined_data[section], list):
                            for i, df in enumerate(combined_data[section]):
                                if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                                    # Make a copy of the original labels to ensure they're not modified
                                    original_labels = df['region_label'].copy()
                                    # Ensure the region labels are preserved exactly as they are
                                    df['region_label'] = original_labels

                                    # Get page number and region labels for debugging
                                    page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else "unknown"
                                    region_labels = df['region_label'].tolist() if 'region_label' in df.columns else []
                                    print(f"[DEBUG] Preserved original region labels for {section}[{i}] from page {page_num} in copy_all_data_to_clipboard: {region_labels}")
                        elif isinstance(combined_data[section], pd.DataFrame) and not combined_data[section].empty and 'region_label' in combined_data[section].columns:
                            # Make a copy of the original labels to ensure they're not modified
                            original_labels = combined_data[section]['region_label'].copy()
                            # Ensure the region labels are preserved exactly as they are
                            combined_data[section]['region_label'] = original_labels

                            # Get page numbers and region labels for debugging
                            page_nums = combined_data[section]['page_number'].unique().tolist() if 'page_number' in combined_data[section].columns else []
                            region_labels = combined_data[section]['region_label'].tolist() if 'region_label' in combined_data[section].columns else []
                            print(f"[DEBUG] Preserved original region labels for {section} from pages {page_nums} in copy_all_data_to_clipboard: {region_labels}")

            # Convert the combined data to text in the same format as sent to invoice2data
            text = invoice_processing_utils.convert_extraction_to_text(combined_data, pdf_path=self.pdf_path)

            # Copy to clipboard
            clipboard = QApplication.clipboard()
            clipboard.setText(text)

            print(f"[DEBUG] Copied combined data from all pages to clipboard in invoice2data format ({len(text.splitlines())} lines)")

            # Show a brief notification
            widget = self.json_tree
            QToolTip.showText(
                widget.mapToGlobal(QPoint(50, 50)),
                "Combined data from all pages copied to clipboard",
                widget,
                QRect(0, 0, 200, 50),
                2000  # Show for 2 seconds
            )
        else:
            # Single page mode - use current data
            # Get the latest cached extraction data directly
            # This ensures we use the most recent data without forcing a new extraction
            current_data = self._get_current_json_data()

            # Convert the data to text in the same format as sent to invoice2data
            text = invoice_processing_utils.convert_extraction_to_text(current_data, pdf_path=self.pdf_path)

            # Copy to clipboard
            clipboard = QApplication.clipboard()
            clipboard.setText(text)

            print(f"[DEBUG] Copied data to clipboard in invoice2data format ({len(text.splitlines())} lines)")

            # Show a brief notification
            widget = self.json_tree
            QToolTip.showText(
                widget.mapToGlobal(QPoint(50, 50)),
                "Data copied to clipboard in invoice2data format",
                widget,
                QRect(0, 0, 200, 50),
                2000  # Show for 2 seconds
            )

    def build_tree_text(self, item, text_list, level):
        """Recursively build text representation of a tree item"""
        if not item:
            return

        # Add this item's text with proper indentation
        indent = '    ' * level
        field = item.text(0)
        value = item.text(1)

        if value.strip():
            text_list.append(f"{indent}{field}: {value}")
        else:
            text_list.append(f"{indent}{field}:")

        # Process all children
        for i in range(item.childCount()):
            child = item.child(i)
            self.build_tree_text(child, text_list, level + 1)

    def expand_item_recursively(self, item):
        """Recursively expand an item and all its children"""
        if not item:
            return

        # Expand this item
        item.setExpanded(True)

        # Recursively expand all children
        for i in range(item.childCount()):
            child = item.child(i)
            self.expand_item_recursively(child)

    def populate_tree_item(self, parent_item, data):
        """Recursively populate a tree item with data"""
        if isinstance(data, dict):
            # Add dictionary items
            for key, value in data.items():
                if value is None or (isinstance(value, float) and np.isnan(value)):
                    continue

                if isinstance(value, (dict, list)):
                    # Create item with children
                    item = QTreeWidgetItem([str(key), ""])
                    parent_item.addChild(item)
                    self.populate_tree_item(item, value)
                else:
                    # Create leaf item
                    item = QTreeWidgetItem([str(key), str(value)])
                    parent_item.addChild(item)

        elif isinstance(data, list):
            # Add list items
            for i, value in enumerate(data):
                if value is None or (isinstance(value, float) and np.isnan(value)):
                    continue

                if isinstance(value, (dict, list)):
                    # Create item with children
                    item = QTreeWidgetItem([f"Item {i+1}", ""])
                    parent_item.addChild(item)
                    self.populate_tree_item(item, value)
                else:
                    # Create leaf item
                    item = QTreeWidgetItem([f"Item {i+1}", str(value)])
                    parent_item.addChild(item)

    def initialize_from_page_configs(self):
        """Initialize regions and column_lines from page_configs for the current page.
        This is used when a multi-page template is first applied to ensure regions are drawn.
        """
        print(f"[DEBUG] initialize_from_page_configs called for page {self.current_page_index + 1}")

        # Ensure regions and column_lines are initialized
        if not hasattr(self, 'regions') or not self.regions:
            self.regions = {'header': [], 'items': [], 'summary': []}

        if not hasattr(self, 'column_lines') or not self.column_lines:
            self.column_lines = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}

        if not isinstance(self.page_configs, list) or len(self.page_configs) == 0:
            print(f"[DEBUG] No page_configs available for initialization")
            return

        # CRITICAL FIX: Use the unified template mapping logic
        try:
            import invoice_processing_utils

            # Create template data for mapping
            template_data = {
                'template_type': 'multi',
                'page_count': len(self.page_configs),
                'config': {
                    'use_middle_page': getattr(self, 'use_middle_page', False),
                    'fixed_page_count': getattr(self, 'fixed_page_count', False)
                },
                'page_regions': [config.get('original_regions', {}) for config in self.page_configs],
                'page_column_lines': [config.get('original_column_lines', {}) for config in self.page_configs]
            }

            # Get the correct template page index using the unified logic
            template_page_index = invoice_processing_utils.get_template_page_for_pdf_page(
                pdf_page_index=self.current_page_index,
                pdf_total_pages=len(self.pdf_document) if self.pdf_document else 1,
                template_data=template_data
            )

            print(f"[DEBUG] UNIFIED MAPPING: PDF page {self.current_page_index + 1} → Template page {template_page_index + 1}")

        except Exception as e:
            print(f"[WARNING] Could not use unified mapping logic: {e}")
            # Fallback to original logic
            template_page_index = self.get_template_page_index(self.current_page_index)
            print(f"[DEBUG] FALLBACK: Using template page {template_page_index + 1 if template_page_index >= 0 else 'combined'} for PDF page {self.current_page_index + 1}")

        # Special case for combined template pages
        if template_page_index == -1 and hasattr(self, 'use_middle_page') and self.use_middle_page:
            print(f"[DEBUG] Special case: 1-page PDF with use_middle_page=True")
            print(f"[DEBUG] Combining first and last template pages for PDF page {self.current_page_index + 1}")

            # Get the first and last template page indices
            first_template_index = 0
            last_template_index = min(2, len(self.page_configs) - 1) if len(self.page_configs) > 2 else len(self.page_configs) - 1

            # Combine regions and column lines from first and last template pages
            combined_regions, combined_column_lines = self.combine_template_pages(first_template_index, last_template_index)
            self.regions = combined_regions
            self.column_lines = combined_column_lines

            # Force a redraw of the PDF label
            if hasattr(self, 'pdf_label'):
                print(f"[DEBUG] Forcing update of pdf_label from initialize_from_page_configs")
                self.pdf_label.update()
            return

        # Check if we have a valid template page
        if template_page_index < len(self.page_configs) and self.page_configs[template_page_index]:
            page_config = self.page_configs[template_page_index]
            print(f"[DEBUG] Using page_config for template page {template_page_index + 1}")

            # Initialize regions from original_regions in page_config
            if 'original_regions' in page_config:
                self.regions = {}
                for section, regions in page_config['original_regions'].items():
                    self.regions[section] = []
                    for i, region in enumerate(regions):
                        # Create QRect from original coordinates
                        rect = QRect(
                            int(region['x']),
                            int(region['y']),
                            int(region['width']),
                            int(region['height'])
                        )

                        # Create region with label
                        titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
                        prefix = titles.get(section, section[0].upper())
                        label = f"{prefix}{i+1}"  # Add region number (1-based)

                        # Store the rect with its label
                        rect_with_label = {
                            'rect': rect,
                            'label': label
                        }
                        self.regions[section].append(rect_with_label)

                    # Update labels for this section to ensure consistency
                    self._update_region_labels(section)

                print(f"[DEBUG] Initialized regions from original_regions in page_config with proper labels")

            # Initialize column_lines from original_column_lines in page_config
            if 'original_column_lines' in page_config:
                self.column_lines = {}
                for section, lines in page_config['original_column_lines'].items():
                    section_enum = RegionType(section)
                    self.column_lines[section_enum] = []
                    for line in lines:
                        # Create column line from original coordinates
                        # Check if line is in the format we expect (list with dictionaries)
                        if isinstance(line, list) and len(line) >= 2:
                            try:
                                # Convert dictionaries to QPoint objects
                                start_point = None
                                end_point = None
                                region_index = None

                                # Handle first point
                                if isinstance(line[0], dict) and 'x' in line[0] and 'y' in line[0]:
                                    start_point = QPoint(int(line[0]['x']), int(line[0]['y']))
                                elif hasattr(line[0], 'x') and hasattr(line[0], 'y'):
                                    # Already a QPoint
                                    start_point = line[0]

                                # Handle second point
                                if isinstance(line[1], dict) and 'x' in line[1] and 'y' in line[1]:
                                    end_point = QPoint(int(line[1]['x']), int(line[1]['y']))
                                elif hasattr(line[1], 'x') and hasattr(line[1], 'y'):
                                    # Already a QPoint
                                    end_point = line[1]

                                # Handle region index if present
                                if len(line) > 2:
                                    region_index = line[2]

                                # Create the line with proper QPoint objects
                                if start_point and end_point:
                                    if region_index is not None:
                                        self.column_lines[section_enum].append((start_point, end_point, region_index))
                                    else:
                                        self.column_lines[section_enum].append((start_point, end_point))
                                    print(f"[DEBUG] Converted column line: {start_point.x()},{start_point.y()} to {end_point.x()},{end_point.y()}")
                            except Exception as e:
                                print(f"[ERROR] Failed to convert column line: {str(e)}")
                                print(f"[ERROR] Line data: {line}")
                        else:
                            # For backward compatibility, try to use the line as is
                            print(f"[WARNING] Using column line in unexpected format: {line}")
                            self.column_lines[section_enum].append(line)
                print(f"[DEBUG] Initialized column_lines from original_column_lines in page_config")

            # Force a redraw of the PDF label
            if hasattr(self, 'pdf_label'):
                print(f"[DEBUG] Forcing update of pdf_label from initialize_from_page_configs")
                self.pdf_label.update()
        else:
            print(f"[DEBUG] No valid template page available for initialization")

    def apply_template(self, template_data):
        """Apply a template to the current PDF

        Args:
            template_data (dict): Template data containing regions, column lines, and other settings
        """
        if not self.pdf_document:
            QMessageBox.warning(self, "No PDF Loaded", "Please load a PDF before applying a template.")
            return

        try:
            print(f"\n[DEBUG] Applying template: {template_data.get('name', 'Unnamed')}")

            # Initialize template application
            self._initialize_template_application(template_data)

            # Configure extraction parameters
            self._configure_extraction_parameters(template_data)

            # Apply template based on type
            is_multi_page = template_data.get('is_multi_page', False)
            if is_multi_page:
                self._apply_multi_page_template(template_data)
            else:
                self._apply_single_page_template(template_data)

            # Finalize template application
            self._finalize_template_application(template_data)

        except Exception as e:
            self._handle_template_application_error(e, template_data)

    def _initialize_template_application(self, template_data):
        """Initialize the template application process"""
        # Reset cached extraction data to ensure we don't show old results
        self._cached_extraction_data = {
            'header': [],
            'items': [],
            'summary': []
        }
        self._last_extraction_state = None

        # Reset the JSON tree to show we're starting fresh
        self.json_tree.clear()
        placeholder_item = QTreeWidgetItem(["Applying template...", "Please wait while the template is being applied..."])
        self.json_tree.addTopLevelItem(placeholder_item)

        # Process events to update the UI
        QApplication.processEvents()

        # Check if this is a multi-page template
        is_multi_page = template_data.get('is_multi_page', False)
        print(f"[DEBUG] Template is multi-page: {is_multi_page}")

        # Set multi-page mode based on template
        self.multi_page_mode = is_multi_page

        # Set middle page and fixed page settings if available
        if 'use_middle_page' in template_data:
            self.use_middle_page = template_data['use_middle_page']
            print(f"[DEBUG] Set use_middle_page={self.use_middle_page}")

        if 'fixed_page_count' in template_data:
            self.fixed_page_count = template_data['fixed_page_count']
            print(f"[DEBUG] Set fixed_page_count={self.fixed_page_count}")

        # Mark the source as template_manager
        self.pdf_loading_source = "template_manager"
        print(f"[DEBUG] Setting PDF loading source to 'template_manager' for template application")

    def _configure_extraction_parameters(self, template_data):
        """Configure extraction parameters from template data"""
        # Set extraction method if available
        if 'extraction_method' in template_data:
            self.current_extraction_method = template_data['extraction_method']
            print(f"[DEBUG] Set extraction method from template: {self.current_extraction_method}")
        else:
            # Default to pypdf_table_extraction if not specified
            self.current_extraction_method = 'pypdf_table_extraction'
            print(f"[DEBUG] Using default extraction method: {self.current_extraction_method}")

        # Set extraction parameters from template config if available
        if 'config' in template_data:
            print(f"[DEBUG] Found config in template data: {template_data['config']}")
            # Check if config contains extraction_params
            if 'extraction_params' in template_data['config']:
                print(f"[DEBUG] Using extraction_params from config")
                self._set_extraction_params_from_config(template_data['config']['extraction_params'])
            else:
                # Use the config directly as extraction parameters
                print(f"[DEBUG] Using config directly as extraction parameters")
                self._set_extraction_params_from_config(template_data['config'])

        # Set extraction parameters if available (direct parameters override config)
        if 'extraction_params' in template_data:
            print(f"[DEBUG] Using direct extraction_params from template")
            self._set_extraction_params_from_direct(template_data['extraction_params'])

    def _set_extraction_params_from_config(self, config_params):
        """Set extraction parameters from template config"""
        import copy
        self.extraction_params = copy.deepcopy(config_params)
        print(f"[DEBUG] Set extraction parameters from template config: {self.extraction_params}")
        self._ensure_extraction_params_structure()
        self._add_custom_params_from_config()

    def _set_extraction_params_from_direct(self, direct_params):
        """Set extraction parameters from direct template parameters"""
        import copy
        self.extraction_params = copy.deepcopy(direct_params)
        print(f"[DEBUG] Set extraction parameters from template direct parameters")
        self._ensure_extraction_params_structure()
        self._add_custom_params_from_extraction()

    def _ensure_extraction_params_structure(self):
        """Ensure extraction parameters have the required structure"""
        # Ensure extraction parameters have the required structure
        if not isinstance(self.extraction_params, dict):
            print(f"[WARNING] Extraction parameters are not a dictionary: {type(self.extraction_params)}")
            self.extraction_params = {
                'header': {'row_tol': 5},
                'items': {'row_tol': 15},
                'summary': {'row_tol': 10},
                'flavor': 'stream',
                'strip_text': '\n'
            }

        # Ensure all section parameters are properly initialized
        for section in ['header', 'items', 'summary']:
            if section not in self.extraction_params:
                self.extraction_params[section] = {}

            # Ensure section parameters are dictionaries
            if not isinstance(self.extraction_params[section], dict):
                self.extraction_params[section] = {}

        # Get global parameters to use as defaults
        global_flavor = self.extraction_params.get('flavor', 'stream')
        global_split_text = self.extraction_params.get('split_text', True)
        global_strip_text = self.extraction_params.get('strip_text', '\n')

        # Ensure each section has the necessary parameters
        for section in ['header', 'items', 'summary']:
            section_params = self.extraction_params[section]

            # Add missing parameters to section if they don't exist
            if 'row_tol' not in section_params:
                section_params['row_tol'] = {'header': 5, 'items': 15, 'summary': 10}[section]

            if 'flavor' not in section_params:
                section_params['flavor'] = global_flavor

            if 'split_text' not in section_params:
                section_params['split_text'] = global_split_text

            if 'strip_text' not in section_params:
                section_params['strip_text'] = global_strip_text

            if 'edge_tol' not in section_params:
                section_params['edge_tol'] = 0.5

        # Ensure flavor is set at global level
        if 'flavor' not in self.extraction_params:
            self.extraction_params['flavor'] = 'stream'

        # Print detailed extraction parameters for debugging
        print(f"[DEBUG] Final extraction parameters structure: {self.extraction_params}")
        for section in ['header', 'items', 'summary']:
            print(f"[DEBUG] {section.capitalize()} extraction parameters: {self.extraction_params.get(section, {})}")

    def _add_custom_params_from_config(self):
        """Add custom parameters from config"""
        # This method can be extended to handle custom parameters from config
        pass

    def _add_custom_params_from_extraction(self):
        """Add custom parameters from extraction_params"""
        # Look for custom parameters in the extraction_params
        for key, value in self.extraction_params.items():
            if key.startswith('custom_param_'):
                print(f"[DEBUG] Found custom parameter: {key} = {value}")

    def _apply_multi_page_template(self, template_data):
        """Apply multi-page template configuration"""
        # For multi-page templates, use page_configs
        if 'page_configs' in template_data:
            self.page_configs = template_data['page_configs']
            print(f"[DEBUG] Set page_configs from template with {len(self.page_configs)} pages")

            # Initialize page_regions and page_column_lines from page_configs
            self.page_regions = {}
            self.page_column_lines = {}

            # Pre-populate page_regions and page_column_lines for ALL pages
            self._pre_populate_all_pages()

            # Show navigation buttons for multi-page PDFs
            self.prev_page_btn.show()
            self.next_page_btn.show()
            self.apply_to_remaining_btn.show()

            # Initialize regions and column lines for the current page
            self.initialize_from_page_configs()
        else:
            print(f"[DEBUG] No page_configs found in template")

    def _pre_populate_all_pages(self):
        """Pre-populate regions and column lines for all PDF pages"""
        print(f"[DEBUG] Pre-populating regions for all {len(self.pdf_document)} PDF pages")

        for pdf_page_idx in range(len(self.pdf_document)):
            try:
                # Use unified mapping logic to determine template page
                import invoice_processing_utils

                template_data_for_mapping = {
                    'template_type': 'multi',
                    'page_count': len(self.page_configs),
                    'config': {
                        'use_middle_page': getattr(self, 'use_middle_page', False),
                        'fixed_page_count': getattr(self, 'fixed_page_count', False)
                    }
                }

                template_page_idx = invoice_processing_utils.get_template_page_for_pdf_page(
                    pdf_page_index=pdf_page_idx,
                    pdf_total_pages=len(self.pdf_document),
                    template_data=template_data_for_mapping
                )

                print(f"[DEBUG] Pre-mapping: PDF page {pdf_page_idx + 1} → Template page {template_page_idx + 1}")

                # Initialize regions and column lines for this PDF page
                self._initialize_page_from_template(pdf_page_idx, template_page_idx)

            except Exception as e:
                print(f"[ERROR] Failed to pre-populate page {pdf_page_idx + 1}: {e}")
                # Initialize empty regions as fallback
                self.page_regions[pdf_page_idx] = {'header': [], 'items': [], 'summary': []}
                self.page_column_lines[pdf_page_idx] = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}

    def _initialize_page_from_template(self, pdf_page_idx, template_page_idx):
        """Initialize a specific PDF page from template page"""
        if template_page_idx < len(self.page_configs):
            page_config = self.page_configs[template_page_idx]

            # Initialize regions
            if 'original_regions' in page_config:
                self.page_regions[pdf_page_idx] = {}
                for section, regions in page_config['original_regions'].items():
                    self.page_regions[pdf_page_idx][section] = []
                    for i, region in enumerate(regions):
                        # Create StandardRegion object - SINGLE FORMAT EVERYWHERE
                        standard_region = self.create_standard_region(
                            int(region['x']), int(region['y']),
                            int(region['width']), int(region['height']),
                            section, i
                        )
                        self.page_regions[pdf_page_idx][section].append(standard_region)

                print(f"[DEBUG] Initialized regions for PDF page {pdf_page_idx + 1} from template page {template_page_idx + 1}")
                for section, regions in self.page_regions[pdf_page_idx].items():
                    print(f"[DEBUG]   {section}: {len(regions)} regions")

            # Initialize column lines
            if 'original_column_lines' in page_config:
                self.page_column_lines[pdf_page_idx] = {}
                for section, lines in page_config['original_column_lines'].items():
                    section_enum = RegionType(section)
                    self.page_column_lines[pdf_page_idx][section_enum] = []
                    for line in lines:
                        if isinstance(line, list) and len(line) >= 2:
                            start_point = QPoint(int(line[0]['x']), int(line[0]['y'])) if isinstance(line[0], dict) else line[0]
                            end_point = QPoint(int(line[1]['x']), int(line[1]['y'])) if isinstance(line[1], dict) else line[1]
                            region_index = line[2] if len(line) > 2 else None
                            if region_index is not None:
                                self.page_column_lines[pdf_page_idx][section_enum].append((start_point, end_point, region_index))
                            else:
                                self.page_column_lines[pdf_page_idx][section_enum].append((start_point, end_point))

    def _apply_single_page_template(self, template_data):
        """Apply single-page template configuration"""
        # For single-page templates, use regions and column_lines directly
        # Check for original_regions first, then fall back to regions
        if 'original_regions' in template_data:
            print(f"[DEBUG] Using original_regions from template")
            self._initialize_single_page_regions(template_data['original_regions'])
        elif 'regions' in template_data:
            print(f"[DEBUG] Using regions from template (converting StandardRegion objects)")
            # Convert StandardRegion objects to dictionary format for initialization
            original_regions = {}
            for section, region_list in template_data['regions'].items():
                original_regions[section] = []
                for region in region_list:
                    from standardized_coordinates import StandardRegion
                    if isinstance(region, StandardRegion):
                        original_regions[section].append({
                            'x': region.rect.x(),
                            'y': region.rect.y(),
                            'width': region.rect.width(),
                            'height': region.rect.height(),
                            'label': region.label
                        })
                        print(f"[DEBUG] Converted StandardRegion {region.label}: UI({region.rect.x()},{region.rect.y()},{region.rect.width()},{region.rect.height()})")
                    else:
                        print(f"[WARNING] Skipping non-StandardRegion object: {type(region)}")
            self._initialize_single_page_regions(original_regions)
        else:
            print(f"[WARNING] No regions found in template data")

        # Initialize column lines from original_column_lines or column_lines
        if 'original_column_lines' in template_data:
            print(f"[DEBUG] Using original_column_lines from template")
            self._initialize_single_page_column_lines(template_data['original_column_lines'])
        elif 'column_lines' in template_data:
            print(f"[DEBUG] Using column_lines from template")
            self._initialize_single_page_column_lines(template_data['column_lines'])
        else:
            print(f"[WARNING] No column lines found in template data")

        # Hide navigation buttons for single-page PDFs
        self.prev_page_btn.hide()
        self.next_page_btn.hide()
        self.apply_to_remaining_btn.hide()

    def _initialize_single_page_regions(self, original_regions):
        """Initialize regions for single-page template"""
        self.regions = {}
        for section, regions in original_regions.items():
            self.regions[section] = []
            for i, region in enumerate(regions):
                # Create StandardRegion object - SINGLE FORMAT EVERYWHERE
                standard_region = self.create_standard_region(
                    int(region['x']), int(region['y']),
                    int(region['width']), int(region['height']),
                    section, i
                )
                self.regions[section].append(standard_region)

            # Update labels for this section to ensure consistency
            self._update_region_labels(section)

        print(f"[DEBUG] Initialized regions from original_regions in template with proper labels")

    def _initialize_single_page_column_lines(self, original_column_lines):
        """Initialize column lines for single-page template"""
        self.column_lines = {}
        for section, lines in original_column_lines.items():
            section_enum = RegionType(section)
            self.column_lines[section_enum] = []
            for line in lines:
                if isinstance(line, list) and len(line) >= 2:
                    try:
                        # Convert dictionaries to QPoint objects
                        start_point = None
                        end_point = None
                        region_index = None

                        # Handle first point
                        if isinstance(line[0], dict) and 'x' in line[0] and 'y' in line[0]:
                            start_point = QPoint(int(line[0]['x']), int(line[0]['y']))
                        elif hasattr(line[0], 'x') and hasattr(line[0], 'y'):
                            start_point = line[0]

                        # Handle second point
                        if isinstance(line[1], dict) and 'x' in line[1] and 'y' in line[1]:
                            end_point = QPoint(int(line[1]['x']), int(line[1]['y']))
                        elif hasattr(line[1], 'x') and hasattr(line[1], 'y'):
                            end_point = line[1]

                        # Handle region index if present
                        if len(line) > 2:
                            region_index = line[2]

                        # Create the line with proper QPoint objects
                        if start_point and end_point:
                            if region_index is not None:
                                self.column_lines[section_enum].append((start_point, end_point, region_index))
                            else:
                                self.column_lines[section_enum].append((start_point, end_point))
                            print(f"[DEBUG] Converted column line: {start_point.x()},{start_point.y()} to {end_point.x()},{end_point.y()}")
                    except Exception as e:
                        print(f"[ERROR] Failed to convert column line: {str(e)}")
                        print(f"[ERROR] Line data: {line}")
                else:
                    # For backward compatibility, try to use the line as is
                    print(f"[WARNING] Using column line in unexpected format: {line}")
                    self.column_lines[section_enum].append(line)
        print(f"[DEBUG] Initialized column_lines from original_column_lines in template")

    def _finalize_template_application(self, template_data):
        """Finalize the template application process"""
        # Update the display
        self.pdf_label.update()

        # Clear extraction cache for this PDF to ensure we don't show old results
        from pdf_extraction_utils import clear_extraction_cache_for_pdf
        clear_extraction_cache_for_pdf(self.pdf_path)

        # Ensure extraction parameters have default values
        self._ensure_default_extraction_params()

        # Print detailed extraction parameters for debugging
        print(f"[DEBUG] Using extraction parameters for extraction: {self.extraction_params}")
        for section in ['header', 'items', 'summary']:
            print(f"[DEBUG] {section.capitalize()} extraction parameters: {self.extraction_params.get(section, {})}")

        # Update the display to show the regions with proper labels
        self.pdf_label.update()

        # Extract data with the new template - always force extraction to ensure results are updated
        self.update_extraction_results(force=True)

        # Show confirmation message
        QMessageBox.information(
            self,
            "Template Applied",
            f"Template '{template_data.get('name', 'Unnamed')}' has been applied successfully."
        )

    def _ensure_default_extraction_params(self):
        """Ensure extraction parameters have default values for each section"""
        if not hasattr(self, 'extraction_params') or not self.extraction_params:
            self.extraction_params = {
                'header': {'row_tol': 5},
                'items': {'row_tol': 15},
                'summary': {'row_tol': 10},
                'flavor': 'stream',
                'strip_text': '\n'
            }
        else:
            # Ensure each section exists in extraction parameters
            for section in ['header', 'items', 'summary']:
                if section not in self.extraction_params:
                    self.extraction_params[section] = {'row_tol': 10}
                elif not isinstance(self.extraction_params[section], dict):
                    self.extraction_params[section] = {'row_tol': 10}

            # Ensure flavor is set
            if 'flavor' not in self.extraction_params:
                self.extraction_params['flavor'] = 'stream'

    def _ensure_extraction_params_structure(self):
        """Ensure extraction parameters have proper structure for each section"""
        if not hasattr(self, 'extraction_params') or not self.extraction_params:
            self._ensure_default_extraction_params()
            return

        # Ensure each section has proper structure
        for section in ['header', 'items', 'summary']:
            if section not in self.extraction_params:
                self.extraction_params[section] = {'row_tol': 10}
            elif not isinstance(self.extraction_params[section], dict):
                self.extraction_params[section] = {'row_tol': 10}

            # Add default parameters to each section
            section_params = self.extraction_params[section]
            if 'flavor' not in section_params:
                section_params['flavor'] = self.extraction_params.get('flavor', 'stream')
            if 'split_text' not in section_params:
                section_params['split_text'] = True
            if 'strip_text' not in section_params:
                section_params['strip_text'] = self.extraction_params.get('strip_text', '\n')
            if 'edge_tol' not in section_params:
                section_params['edge_tol'] = 0.5

        print(f"[DEBUG] Ensured extraction params structure: {self.extraction_params}")

    def _handle_template_application_error(self, error, template_data):
        """Handle errors during template application"""
        print(f"Error applying template: {str(error)}")
        import traceback
        traceback.print_exc()

        # Show error message
        QMessageBox.critical(
            self,
            "Error Applying Template",
            f"An error occurred while applying the template: {str(error)}"
        )

    def save_template(self):
        """Save the current configuration as a template"""
        try:
            # Get template name from user
            template_name, ok = QInputDialog.getText(
                self, "Save Template", "Enter template name:"
            )

            if not ok or not template_name:
                return

            # Prepare template data
            template_data = {
                'name': template_name,
                'extraction_params': self.extraction_params
            }

            # Get the invoice2data template JSON from the template_preview
            print(f"\n[DEBUG] Getting invoice2data template JSON for save_template")
            json_template = None

            # Try to get JSON template from template_preview if available
            print(f"[DEBUG] Has template_preview: {hasattr(self, 'template_preview')}")
            if hasattr(self, 'template_preview') and self.template_preview:
                try:
                    json_template_text = self.template_preview.toPlainText()
                    print(f"[DEBUG] invoice2data template text length: {len(json_template_text)}")
                    print(f"[DEBUG] invoice2data template text (first 100 chars): {json_template_text[:100]}...")

                    if json_template_text.strip():
                        json_template = json.loads(json_template_text)
                        print(f"[DEBUG] invoice2data template parsed successfully")
                        print(f"[DEBUG] invoice2data template type: {type(json_template)}")
                        if isinstance(json_template, dict):
                            print(f"[DEBUG] invoice2data template keys: {list(json_template.keys())}")
                            # Add JSON template to template data
                            template_data['json_template'] = json_template
                    else:
                        print(f"[DEBUG] invoice2data template text is empty")
                except json.JSONDecodeError as e:
                    print(f"[WARNING] Invalid invoice2data template: {str(e)}")
                except Exception as e:
                    print(f"[ERROR] Error processing invoice2data template: {str(e)}")
                    import traceback
                    traceback.print_exc()
            else:
                print(f"[DEBUG] No template_preview found, no invoice2data template will be saved")

            # Add multi-page configuration if applicable
            if self.multi_page_mode:
                template_data['is_multi_page'] = True
                # Removed use_middle_page and fixed_page_count - using simplified page-wise approach

                # Store page configurations
                template_data['page_configs'] = []

                # Store regions and column lines for each page
                for page_idx in range(len(self.pdf_document)):
                    page_config = {}

                    # Get regions for this page
                    if isinstance(self.page_regions, dict) and page_idx in self.page_regions:
                        page_config['regions'] = self.page_regions[page_idx]

                        # Store original (unscaled) regions for template application
                        original_regions = {}
                        for section, region_list in self.page_regions[page_idx].items():
                            original_regions[section] = []
                            for region in region_list:
                                # Use standardized coordinate system - NO backward compatibility
                                from standardized_coordinates import StandardRegion
                                if isinstance(region, StandardRegion):
                                    original_regions[section].append({
                                        'x': region.rect.x(),
                                        'y': region.rect.y(),
                                        'width': region.rect.width(),
                                        'height': region.rect.height(),
                                        'label': region.label
                                    })
                                else:
                                    print(f"[ERROR] Invalid region type in template save: expected StandardRegion, got {type(region)}")
                        page_config['original_regions'] = original_regions
                    else:
                        page_config['regions'] = {'header': [], 'items': [], 'summary': []}
                        page_config['original_regions'] = {'header': [], 'items': [], 'summary': []}

                    # Get column lines for this page
                    if isinstance(self.page_column_lines, dict) and page_idx in self.page_column_lines:
                        # Convert column lines to serializable format
                        column_lines_dict = {}
                        for section, lines in self.page_column_lines[page_idx].items():
                            section_name = section.value if hasattr(section, 'value') else section
                            column_lines_dict[section_name] = lines
                        page_config['column_lines'] = column_lines_dict

                        # Store original (unscaled) column lines for template application
                        page_config['original_column_lines'] = column_lines_dict
                    else:
                        page_config['column_lines'] = {}
                        page_config['original_column_lines'] = {}

                    template_data['page_configs'].append(page_config)
            else:
                # Single-page mode
                template_data['is_multi_page'] = False

                # Store regions
                template_data['regions'] = self.regions

                # Store original (unscaled) regions for template application
                original_regions = {}
                for section, region_list in self.regions.items():
                    original_regions[section] = []
                    for region in region_list:
                        # Use standardized coordinate system - NO backward compatibility
                        from standardized_coordinates import StandardRegion
                        if isinstance(region, StandardRegion):
                            original_regions[section].append({
                                'x': region.rect.x(),
                                'y': region.rect.y(),
                                'width': region.rect.width(),
                                'height': region.rect.height(),
                                'label': region.label
                            })
                        else:
                            print(f"[ERROR] Invalid region type in template save: expected StandardRegion, got {type(region)}")
                template_data['original_regions'] = original_regions

                # Store column lines
                # Convert column lines to serializable format
                column_lines_dict = {}
                for section, lines in self.column_lines.items():
                    section_name = section.value if hasattr(section, 'value') else section
                    column_lines_dict[section_name] = lines
                template_data['column_lines'] = column_lines_dict

                # Store original (unscaled) column lines for template application
                template_data['original_column_lines'] = column_lines_dict

            # Actually save the template to the database
            try:
                # Get description from user
                description, ok = QInputDialog.getText(
                    self, "Template Description", "Enter template description (optional):"
                )
                if not ok:
                    description = ""

                # Collect actual regions and convert to dual coordinate format
                print(f"[DEBUG] Collecting regions for template save...")

                # Get current regions and column lines from the application state
                current_regions = self.get_current_regions_for_save()
                current_column_lines = self.get_current_column_lines_for_save()

                print(f"[DEBUG] Collected regions: {list(current_regions.keys())}")
                for section, region_list in current_regions.items():
                    print(f"[DEBUG]   {section}: {len(region_list)} regions")

                print(f"[DEBUG] Collected column lines: {list(current_column_lines.keys())}")
                for section, line_list in current_column_lines.items():
                    print(f"[DEBUG]   {section}: {len(line_list)} column lines")

                # Convert regions to dual coordinate format
                dual_regions, dual_column_lines = self.convert_regions_to_dual_coordinates(
                    current_regions, current_column_lines
                )

                template_id = self.db.save_template(
                    name=template_name,
                    description=description,
                    config=template_data.get('extraction_params', {}),
                    template_type='multi' if template_data.get('is_multi_page', False) else 'single',
                    page_count=len(template_data.get('page_configs', [])) if template_data.get('is_multi_page', False) else 1,
                    json_template=template_data.get('json_template', {}),
                    # Dual coordinate fields with actual data
                    drawing_regions=dual_regions,
                    drawing_column_lines=dual_column_lines,
                    extraction_regions=dual_regions,  # Same data, different usage
                    extraction_column_lines=dual_column_lines,  # Same data, different usage
                    extraction_method=getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
                )

                print(f"[DEBUG] Template saved to database with ID: {template_id}")

                # Emit signal to notify other components
                self.save_template_signal.emit()

                # Show confirmation using factory
                UIMessageFactory.show_info(
                    self,
                    "Template Saved",
                    f"Template '{template_name}' has been saved successfully with ID {template_id}."
                )

            except Exception as db_error:
                print(f"[ERROR] Failed to save template to database: {db_error}")
                import traceback
                traceback.print_exc()

                # Show error message using factory
                UIMessageFactory.show_error(
                    self,
                    "Database Error",
                    f"Failed to save template to database: {str(db_error)}"
                )
                return

        except Exception as e:
            print(f"Error saving template: {str(e)}")
            import traceback
            traceback.print_exc()

            # Show error message using factory
            UIMessageFactory.show_error(
                self,
                "Error Saving Template",
                f"An error occurred while saving the template: {str(e)}"
            )

    def get_current_regions_for_save(self):
        """Get current regions from the application state for saving"""
        try:
            if self.multi_page_mode:
                # For multi-page mode, collect regions from all pages
                all_regions = {'header': [], 'items': [], 'summary': []}
                for page_idx, page_regions in self.page_regions.items():
                    for section, region_list in page_regions.items():
                        all_regions[section].extend(region_list)
                return all_regions
            else:
                # For single-page mode, use current regions
                # Check if we have regions from the main PDF processor
                if hasattr(self, 'parent') and hasattr(self.parent, 'pdf_processor'):
                    main_regions = getattr(self.parent.pdf_processor, 'regions', {})
                    if main_regions:
                        print(f"[DEBUG] Using regions from main PDF processor: {list(main_regions.keys())}")
                        return main_regions.copy()

                # Fallback to local regions
                print(f"[DEBUG] Using local regions: {list(self.regions.keys())}")
                return self.regions.copy()
        except Exception as e:
            print(f"[ERROR] Error getting current regions: {e}")
            import traceback
            traceback.print_exc()
            return {'header': [], 'items': [], 'summary': []}

    def get_current_column_lines_for_save(self):
        """Get current column lines from the application state for saving"""
        try:
            if self.multi_page_mode:
                # For multi-page mode, collect column lines from all pages
                all_column_lines = {'header': [], 'items': [], 'summary': []}
                for page_idx, page_column_lines in self.page_column_lines.items():
                    for section, line_list in page_column_lines.items():
                        # Convert RegionType enum to string if needed
                        section_key = section.value if hasattr(section, 'value') else str(section).lower()
                        if section_key in all_column_lines:
                            all_column_lines[section_key].extend(line_list)
                return all_column_lines
            else:
                # For single-page mode, use current column lines
                # Check if we have column lines from the main PDF processor
                main_column_lines = {}
                if hasattr(self, 'parent') and hasattr(self.parent, 'pdf_processor'):
                    main_column_lines = getattr(self.parent.pdf_processor, 'column_lines', {})
                    if main_column_lines:
                        print(f"[DEBUG] Using column lines from main PDF processor: {list(main_column_lines.keys())}")

                # Use main column lines if available, otherwise use local
                source_column_lines = main_column_lines if main_column_lines else self.column_lines

                # Convert RegionType keys to string keys
                converted_column_lines = {'header': [], 'items': [], 'summary': []}
                for section, line_list in source_column_lines.items():
                    # Convert RegionType enum to string if needed
                    section_key = section.value if hasattr(section, 'value') else str(section).lower()
                    if section_key in converted_column_lines:
                        converted_column_lines[section_key] = line_list.copy()
                    else:
                        print(f"[WARNING] Unknown section key: {section} ({type(section)})")

                print(f"[DEBUG] Converted column lines keys: {list(converted_column_lines.keys())}")
                for section, lines in converted_column_lines.items():
                    print(f"[DEBUG]   {section}: {len(lines)} lines")

                return converted_column_lines
        except Exception as e:
            print(f"[ERROR] Error getting current column lines: {e}")
            import traceback
            traceback.print_exc()
            return {'header': [], 'items': [], 'summary': []}

    def convert_regions_to_dual_coordinates(self, regions, column_lines):
        """Convert regions and column lines to dual coordinate format"""
        try:
            from dual_coordinate_storage import DualCoordinateRegion, DualCoordinateColumnLine

            # Get scale factors for coordinate conversion
            scale_factors = get_scale_factors(self.pdf_path, self.current_page_index)
            scale_x = scale_factors['scale_x']
            scale_y = scale_factors['scale_y']
            page_height = scale_factors['page_height']

            print(f"[DEBUG] Using scale factors: scale_x={scale_x}, scale_y={scale_y}, page_height={page_height}")

            # Convert regions to dual coordinate format
            dual_regions = {'header': [], 'items': [], 'summary': []}
            for section, region_list in regions.items():
                for region in region_list:
                    try:
                        # Handle StandardRegion objects
                        if hasattr(region, 'rect') and hasattr(region, 'label'):
                            rect = region.rect
                            label = region.label

                            # Create dual coordinate region
                            dual_region = DualCoordinateRegion.from_qrect(
                                rect, label, scale_x, scale_y, page_height
                            )
                            dual_regions[section].append(dual_region)
                            print(f"[DEBUG] Converted region {label} to dual coordinates")

                        elif isinstance(region, dict) and 'rect' in region:
                            # Handle dictionary format
                            rect = region['rect']
                            label = region.get('label', f"{section[0].upper()}1")

                            dual_region = DualCoordinateRegion.from_qrect(
                                rect, label, scale_x, scale_y, page_height
                            )
                            dual_regions[section].append(dual_region)
                            print(f"[DEBUG] Converted dict region {label} to dual coordinates")

                        else:
                            print(f"[WARNING] Unknown region format in {section}: {type(region)}")

                    except Exception as e:
                        print(f"[ERROR] Error converting region in {section}: {e}")
                        continue

            # Convert column lines to dual coordinate format
            dual_column_lines = {'header': [], 'items': [], 'summary': []}
            for section, line_list in column_lines.items():
                print(f"[DEBUG] Converting column lines for section: {section} ({len(line_list)} lines)")
                for i, line in enumerate(line_list):
                    try:
                        print(f"[DEBUG] Processing column line {i}: {type(line)}")

                        # Handle different column line formats
                        if isinstance(line, (list, tuple)) and len(line) >= 2:
                            start_point = line[0]
                            end_point = line[1]

                            print(f"[DEBUG] Start point type: {type(start_point)}, End point type: {type(end_point)}")

                            # Extract coordinates
                            if hasattr(start_point, 'x') and hasattr(start_point, 'y'):
                                start_x, start_y = start_point.x(), start_point.y()
                                end_x, end_y = end_point.x(), end_point.y()
                                print(f"[DEBUG] QPoint coordinates: ({start_x}, {start_y}) -> ({end_x}, {end_y})")
                            elif isinstance(start_point, dict):
                                start_x, start_y = start_point['x'], start_point['y']
                                end_x, end_y = end_point['x'], end_point['y']
                                print(f"[DEBUG] Dict coordinates: ({start_x}, {start_y}) -> ({end_x}, {end_y})")
                            else:
                                print(f"[WARNING] Unknown column line format in {section}: {type(start_point)}")
                                print(f"[DEBUG] Line data: {line}")
                                continue

                            # Create dual coordinate column line
                            label = f"C{len(dual_column_lines[section]) + 1}"
                            dual_line = DualCoordinateColumnLine.from_ui_input(
                                start_x, start_y, end_x, end_y, scale_x, scale_y, page_height, label
                            )
                            dual_column_lines[section].append(dual_line)
                            print(f"[DEBUG] Converted column line {label} to dual coordinates")
                        else:
                            print(f"[WARNING] Invalid column line format in {section}: {type(line)} with length {len(line) if hasattr(line, '__len__') else 'N/A'}")
                            print(f"[DEBUG] Line data: {line}")

                    except Exception as e:
                        print(f"[ERROR] Error converting column line in {section}: {e}")
                        import traceback
                        traceback.print_exc()
                        continue

            return dual_regions, dual_column_lines

        except Exception as e:
            print(f"[ERROR] Error converting to dual coordinates: {e}")
            return {'header': [], 'items': [], 'summary': []}, {'header': [], 'items': [], 'summary': []}

    # Complex page mapping methods removed - using simplified page-wise approach
    # Page mapping features will be implemented later as a separate enhancement

    def extract_specific_region(self, section_type, region_index):
        """Extract data for a specific region after a column line is drawn

        Args:
            section_type (str): The section type ('header', 'items', or 'summary')
            region_index (int): The index of the region within the section
        """
        if not self.pdf_document or not self.pdf_path:
            return

        try:
            print(f"\n[DEBUG] Extracting specific region: {section_type}, index {region_index}, page {self.current_page_index + 1}")

            # Get the appropriate regions dictionary based on mode
            regions_dict = self.page_regions.get(self.current_page_index, {}) if self.multi_page_mode else self.regions

            # Validate that the section type exists
            if section_type not in regions_dict:
                print(f"[DEBUG] Section type {section_type} not found in regions for page {self.current_page_index + 1}")
                # Initialize empty list for this section if it doesn't exist
                if self.multi_page_mode:
                    if self.current_page_index not in self.page_regions:
                        self.page_regions[self.current_page_index] = {'header': [], 'items': [], 'summary': []}

                    # Initialize the section type in both dictionaries
                    self.page_regions[self.current_page_index][section_type] = []
                    self.regions[section_type] = []

                    print(f"[DEBUG] Initialized empty {section_type} list for page {self.current_page_index + 1}")
                else:
                    self.regions[section_type] = []

                # Now that we've initialized the section, we can continue with an empty list
                regions_dict = self.page_regions.get(self.current_page_index, {}) if self.multi_page_mode else self.regions
                regions = regions_dict.get(section_type, [])

            # Get the regions for this section
            regions = regions_dict.get(section_type, [])

            # Validate that the region index is valid
            if region_index >= len(regions):
                print(f"[DEBUG] Region index {region_index} is out of bounds for section {section_type}")
                return

            # Get the region rect and label
            region_rect = None
            region_label = None

            # Get the region item
            region_item = regions[region_index]

            # Use standardized coordinate system - NO backward compatibility
            from standardized_coordinates import StandardRegion
            if isinstance(region_item, StandardRegion):
                region_rect = region_item.rect
                region_label = region_item.label
            else:
                print(f"[ERROR] Invalid region type in extract_specific_region: expected StandardRegion, got {type(region_item)}")
                return

            if not region_rect:
                print(f"[DEBUG] Region not found: {section_type}, index {region_index}")
                return

            # Ensure region_rect is a QRect object
            if not hasattr(region_rect, 'x') or not callable(region_rect.x):
                print(f"[DEBUG] Invalid region_rect type: {type(region_rect)}")
                return

            # StandardRegion should always have a label - this is just a safety check
            if not region_label:
                print(f"[ERROR] StandardRegion missing label - this should not happen!")
                titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
                prefix = titles.get(section_type, section_type[0].upper())
                region_label = f"{prefix}{region_index + 1}"  # Add region number (1-based)

            # Get column lines for this region
            column_lines = []
            region_type = RegionType(section_type) if section_type in ['header', 'items', 'summary'] else None

            if region_type:
                # Get the appropriate column lines dictionary based on mode
                column_lines_dict = self.page_column_lines.get(self.current_page_index, {}) if self.multi_page_mode else self.column_lines

                # Try to get column lines using enum key
                lines = column_lines_dict.get(region_type, [])

                # If no lines found with enum key, try string key
                if not lines and section_type in column_lines_dict:
                    lines = column_lines_dict.get(section_type, [])

                print(f"[DEBUG] Found {len(lines)} column lines for region type {region_type}")

                # Process each column line
                for line in lines:
                    # Check if the line has a region index (format: (start_point, end_point, region_index))
                    if len(line) > 2 and isinstance(line[2], int):
                        # Only add if the region index matches
                        if line[2] == region_index:
                            column_lines.append(line[0].x())
                            print(f"[DEBUG] Added column at x={line[0].x()} for region {region_index} (matched region index)")
                    else:
                        # For lines without region index, add to all regions in multipage mode
                        # or only to region 0 in single-page mode
                        if self.multi_page_mode or region_index == 0:
                            column_lines.append(line[0].x())
                            print(f"[DEBUG] Added column at x={line[0].x()} for region {region_index} (no region index)")

            # Sort column lines by x-coordinate
            column_lines.sort()

            # Get PDF page dimensions and scale factors
            scale_factors = get_scale_factors(self.pdf_path, self.current_page_index)
            scale_x = scale_factors['scale_x']
            scale_y = scale_factors['scale_y']
            page_height = scale_factors['page_height']

            # Convert region rect to PDF coordinates
            x1 = region_rect.x() * scale_x
            y1 = page_height - (region_rect.y() * scale_y)  # Convert top y1 to bottom y1
            x2 = (region_rect.x() + region_rect.width()) * scale_x
            y2 = page_height - ((region_rect.y() + region_rect.height()) * scale_y)  # Convert bottom y2 to bottom y2

            # Create table area
            table_area = [x1, y1, x2, y2]

            # Convert column lines to PDF coordinates and format as comma-separated string
            if column_lines:
                pdf_columns = [x * scale_x for x in column_lines]
                pdf_columns_str = ','.join(map(str, pdf_columns))
                print(f"[DEBUG] Table area: {table_area}")
                print(f"[DEBUG] Columns: {pdf_columns_str}")
            else:
                pdf_columns_str = None
                print(f"[DEBUG] Table area: {table_area}")
                print(f"[DEBUG] No columns for this region")

            # Extract data for this specific region - use params directly without defaults
            # Print the current extraction parameters for debugging
            print(f"[DEBUG] Current extraction parameters: {self.extraction_params}")

            # Get section-specific parameters
            section_params = self.extraction_params.get(section_type, {}).copy()
            print(f"[DEBUG] Section-specific parameters for {section_type}: {section_params}")

            # Check if row_tol exists in section parameters
            if 'row_tol' in section_params:
                print(f"[DEBUG] Found row_tol for {section_type}: {section_params['row_tol']}")
            else:
                print(f"[DEBUG] No row_tol found for {section_type}")

            # Prepare extraction parameters - include all necessary parameters
            extraction_params = {
                'header': {},
                'items': {},
                'summary': {},
                'pages': str(self.current_page_index + 1)  # 1-based page number
            }

            # Make sure the pages parameter is set in all sections
            extraction_params['header']['pages'] = str(self.current_page_index + 1)
            extraction_params['items']['pages'] = str(self.current_page_index + 1)
            extraction_params['summary']['pages'] = str(self.current_page_index + 1)

            # Get global parameters to use as defaults for section-specific parameters
            global_flavor = self.extraction_params.get('flavor', 'stream')
            global_split_text = self.extraction_params.get('split_text', True)
            global_strip_text = self.extraction_params.get('strip_text', '\n')

            # Ensure section parameters have all necessary values
            if not section_params:
                section_params = {}

            # Add missing parameters to section if they don't exist
            if 'row_tol' not in section_params:
                if section_type == 'header':
                    section_params['row_tol'] = 5
                elif section_type == 'items':
                    section_params['row_tol'] = 15
                else:  # summary
                    section_params['row_tol'] = 10

            if 'flavor' not in section_params:
                section_params['flavor'] = global_flavor

            if 'split_text' not in section_params:
                section_params['split_text'] = global_split_text

            if 'strip_text' not in section_params:
                section_params['strip_text'] = global_strip_text

            # Add edge_tol if not present
            if 'edge_tol' not in section_params:
                section_params['edge_tol'] = 0.5

            # Check if we have columns and need to adjust the flavor
            if pdf_columns_str:
                # Get the flavor from section parameters
                section_flavor = section_params.get('flavor', 'stream')

                if section_flavor == 'lattice':
                    print(f"[WARNING] Columns are defined but flavor is set to 'lattice'. Changing to 'stream' flavor.")
                    # Update the section flavor to 'stream'
                    section_params['flavor'] = 'stream'

            # Update the extraction_params with the enhanced section_params
            extraction_params[section_type] = section_params

            # Print the section parameters for debugging
            print(f"[DEBUG] Using section parameters for {section_type}: {extraction_params[section_type]}")

            # Print the final extraction parameters for debugging
            print(f"[DEBUG] Final extraction parameters for extract_table: {extraction_params}")
            print(f"[DEBUG] Section parameters for {section_type}: {extraction_params.get(section_type, {})}")

            # Prepare additional parameters from custom parameters
            additional_params = {}

            # Check for section-specific custom parameters
            for i in range(1, 10):  # Support up to 9 custom parameters
                param_name_key = f'{section_type}_custom_param_{i}_name'
                param_value_key = f'{section_type}_custom_param_{i}_value'

                if param_name_key in self.extraction_params and param_value_key in self.extraction_params:
                    param_name = self.extraction_params[param_name_key]
                    param_value = self.extraction_params[param_value_key]

                    if param_name and param_value is not None:
                        additional_params[param_name] = param_value
                        print(f"[DEBUG] Using section-specific custom parameter for {section_type}: {param_name} = {param_value}")

            # Also check for global custom parameters
            for i in range(1, 10):  # Support up to 9 custom parameters
                param_name_key = f'custom_param_{i}_name'
                param_value_key = f'custom_param_{i}_value'

                if param_name_key in self.extraction_params and param_value_key in self.extraction_params:
                    param_name = self.extraction_params[param_name_key]
                    param_value = self.extraction_params[param_value_key]

                    if param_name and param_value is not None:
                        # Only add if not already added as a section-specific parameter
                        if param_name not in additional_params:
                            additional_params[param_name] = param_value
                            print(f"[DEBUG] Using global custom parameter: {param_name} = {param_value}")

            # Check for direct custom parameters (without name/value pairs)
            for key, value in self.extraction_params.items():
                if key.startswith('custom_param_') and not key.endswith('_name') and not key.endswith('_value'):
                    # This is a direct custom parameter
                    additional_params[key] = value
                    print(f"[DEBUG] Using direct custom parameter: {key} = {value}")

            # Print extraction parameters before calling extract_table
            print(f"[DEBUG] Calling extract_table for {section_type} with enhanced extraction_params: {extraction_params}")
            print(f"[DEBUG] Enhanced {section_type.capitalize()} section parameters: {extraction_params.get(section_type, {})}")
            print(f"[DEBUG] Additional parameters: {additional_params}")

            # Extract table with additional parameters
            # Use the enhanced extraction_params we just created
            print(f"[DEBUG] Ignoring cached data and forcing re-extraction for specific region")
            print(f"[DEBUG] Using region label: {region_label}")

            # Clear the extraction cache for this section, but preserve the multi-page cache
            # This is critical to prevent regions with the same label on different pages from being replaced
            clear_extraction_cache_for_section(self.pdf_path, self.current_page_index + 1, section_type, preserve_multipage=True)

            # Extract using multi-method extraction
            extraction_method = getattr(self, 'current_extraction_method', 'pypdf_table_extraction')
            df = extract_with_method(
                pdf_path=self.pdf_path,
                extraction_method=extraction_method,
                page_number=self.current_page_index + 1,  # 1-based page number
                table_areas=[table_area],
                columns_list=[pdf_columns_str] if pdf_columns_str else None,
                section_type=section_type,
                extraction_params=extraction_params,  # Use the enhanced extraction_params we created
                use_cache=False  # Disable caching to force re-extraction
            )

            if df is not None and not df.empty:
                print(f"[DEBUG] Successfully extracted data for {section_type} region {region_index}")
                print(f"[DEBUG] Extracted data:\n{df}")

                # Add page number to the DataFrame if it doesn't already have one
                if 'page_number' not in df.columns:
                    df['page_number'] = self.current_page_index + 1  # 1-based page number
                if '_page_number' not in df.columns:
                    df['_page_number'] = self.current_page_index + 1  # 1-based page number

                # Add region_label column if it doesn't exist
                if 'region_label' not in df.columns:
                    # Get the region label from the region dictionary first
                    # This ensures we use the original label that was assigned when the region was drawn
                    stored_region_label = None

                    # Try to get the existing label from the region dictionary
                    if self.multi_page_mode:
                        if self.current_page_index in self.page_regions and section_type in self.page_regions[self.current_page_index]:
                            if region_index < len(self.page_regions[self.current_page_index][section_type]):
                                region_item = self.page_regions[self.current_page_index][section_type][region_index]
                                if isinstance(region_item, dict) and 'label' in region_item:
                                    stored_region_label = region_item['label']
                                    print(f"[DEBUG] Found existing region label in page_regions: {stored_region_label}")
                    else:
                        if section_type in self.regions:
                            if region_index < len(self.regions[section_type]):
                                region_item = self.regions[section_type][region_index]
                                if isinstance(region_item, dict) and 'label' in region_item:
                                    stored_region_label = region_item['label']
                                    print(f"[DEBUG] Found existing region label in regions: {stored_region_label}")

                    # Use the stored label if available, otherwise use the passed label, or create a default one
                    if stored_region_label:
                        region_label = stored_region_label
                        print(f"[DEBUG] Using stored region label: {region_label}")
                    elif not region_label:
                        titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
                        prefix = titles.get(section_type, section_type[0].upper())
                        region_label = f"{prefix}{region_index + 1}"  # Add region number (1-based)
                        print(f"[DEBUG] Created default region label: {region_label}")

                    # Store the region label for future reference
                    # Update the region item with the label if it's a dictionary
                    if self.multi_page_mode:
                        if self.current_page_index in self.page_regions and section_type in self.page_regions[self.current_page_index]:
                            if region_index < len(self.page_regions[self.current_page_index][section_type]):
                                region_item = self.page_regions[self.current_page_index][section_type][region_index]
                                if isinstance(region_item, dict) and 'rect' in region_item:
                                    # Only update if the label doesn't exist or is different
                                    if 'label' not in region_item or region_item['label'] != region_label:
                                        region_item['label'] = region_label
                                        print(f"[DEBUG] Updated region label in page_regions: {region_label}")
                    else:
                        if section_type in self.regions:
                            if region_index < len(self.regions[section_type]):
                                region_item = self.regions[section_type][region_index]
                                if isinstance(region_item, dict) and 'rect' in region_item:
                                    # Only update if the label doesn't exist or is different
                                    if 'label' not in region_item or region_item['label'] != region_label:
                                        region_item['label'] = region_label
                                        print(f"[DEBUG] Updated region label in regions: {region_label}")

                    # Add row numbers to each label with page numbers to differentiate regions with the same name on different pages
                    # Use the exact region label from the UI without modification
                    # IMPORTANT: Preserve the original region label (H1, H2, etc.) exactly as it is
                    df['region_label'] = [f"{region_label}_R{j+1}_P{self.current_page_index+1}"
                                        for j in range(len(df))]

                    # Log the region labels for debugging
                    print(f"[DEBUG] Added region labels using {region_label} to extracted DataFrame")
                    print(f"[DEBUG] Created region labels: {df['region_label'].tolist()}")
                    print(f"[DEBUG] Using exact region label from UI: {region_label} (preserving region number)")

                    # Analyze the multi-page labels to ensure they have the correct format
                    labels = df['region_label'].tolist()
                    if labels:
                        print(f"[DEBUG] Found {len(labels)} multi-page labels with page info")
                        for label in labels[:5]:  # Show first 5 labels for debugging
                            parts = label.split('_')
                            if len(parts) >= 3:
                                region_part = parts[0]  # H1, H2, etc.
                                row_part = parts[1]     # R1, R2, etc.
                                page_part = parts[2]    # P1, P2, etc.
                                print(f"[DEBUG] Label {label}: Region={region_part}, Row={row_part}, Page={page_part}")

                    print(f"[DEBUG] Preserving ALL regions from all pages without duplicate checking")
                    print(f"[DEBUG] Ensuring data from page {self.current_page_index+1} is preserved in multi-page mode")
                else:
                    # DO NOT modify region labels - preserve them exactly as they are
                    # Just verify that the labels are present and log them
                    labels = df['region_label'].tolist()
                    print(f"[DEBUG] Preserving original region labels exactly as is: {labels}")
                    print(f"[DEBUG] Not modifying existing region labels to ensure consistency")

                    # CRITICAL: Ensure the original region labels are preserved exactly as they are
                    # This is the key fix to prevent region labels from being changed

                    # Track that these region labels have been set
                    if self.current_page_index not in self._region_labels_set:
                        self._region_labels_set[self.current_page_index] = {}
                    if section_type not in self._region_labels_set[self.current_page_index]:
                        self._region_labels_set[self.current_page_index][section_type] = {}

                    # Mark each label as set
                    for i, label in enumerate(labels):
                        self._region_labels_set[self.current_page_index][section_type][i] = True

                # Get the current data from the cached extraction data
                current_data = self._get_current_json_data()

                print(f"[DEBUG] Current data before update: {list(current_data.keys())}")
                print(f"[DEBUG] Updating {section_type} region {region_index} with new data")

                # Update the specific section with the new data
                if section_type == 'header':
                    # Convert header to a list if it's not already
                    if not isinstance(current_data['header'], list):
                        if current_data['header'] is not None:
                            # Convert single DataFrame to list
                            current_data['header'] = [current_data['header']]
                        else:
                            # Initialize empty list
                            current_data['header'] = []

                    # Update or append the specific region
                    if region_index < len(current_data['header']):
                        current_data['header'][region_index] = df
                        print(f"[DEBUG] Updated existing header region {region_index}")
                    else:
                        # Fill any gaps with None
                        while len(current_data['header']) < region_index:
                            current_data['header'].append(None)
                        current_data['header'].append(df)
                        print(f"[DEBUG] Added new header region {region_index}")
                elif section_type == 'items':
                    # Make sure items is a list
                    if not isinstance(current_data['items'], list):
                        if current_data['items'] is not None:
                            # Convert single DataFrame to list
                            current_data['items'] = [current_data['items']]
                        else:
                            # Initialize empty list
                            current_data['items'] = []

                    # Update or append the specific region
                    if region_index < len(current_data['items']):
                        current_data['items'][region_index] = df
                        print(f"[DEBUG] Updated existing items region {region_index}")
                    else:
                        # Fill any gaps with None
                        while len(current_data['items']) < region_index:
                            current_data['items'].append(None)
                        current_data['items'].append(df)
                        print(f"[DEBUG] Added new items region {region_index}")
                elif section_type == 'summary':
                    # Convert summary to a list if it's not already
                    if not isinstance(current_data['summary'], list):
                        if current_data['summary'] is not None:
                            # Convert single DataFrame to list
                            current_data['summary'] = [current_data['summary']]
                        else:
                            # Initialize empty list
                            current_data['summary'] = []

                    # Update or append the specific region
                    if region_index < len(current_data['summary']):
                        current_data['summary'][region_index] = df
                        print(f"[DEBUG] Updated existing summary region {region_index}")
                    else:
                        # Fill any gaps with None
                        while len(current_data['summary']) < region_index:
                            current_data['summary'].append(None)
                        current_data['summary'].append(df)
                        print(f"[DEBUG] Added new summary region {region_index}")

                print(f"[DEBUG] Current data after update: {list(current_data.keys())}")
                if 'items' in current_data and isinstance(current_data['items'], list):
                    print(f"[DEBUG] Items count: {len(current_data['items'])}")

                # Store the extracted data in _all_pages_data
                if not hasattr(self, '_all_pages_data'):
                    print(f"[DEBUG] _all_pages_data attribute not found, initializing in extract_specific_region")
                    self._all_pages_data = [None] * len(self.pdf_document)
                elif not self._all_pages_data:
                    print(f"[DEBUG] _all_pages_data is empty, initializing in extract_specific_region")
                    self._all_pages_data = [None] * len(self.pdf_document)
                elif len(self._all_pages_data) != len(self.pdf_document):
                    print(f"[DEBUG] _all_pages_data length mismatch, reinitializing in extract_specific_region")
                    self._all_pages_data = [None] * len(self.pdf_document)

                # Update the data for the current page
                if self.current_page_index < len(self._all_pages_data):
                    self._all_pages_data[self.current_page_index] = current_data
                    print(f"[DEBUG] Stored extraction data for page {self.current_page_index + 1} after region extraction")

                    # Print summary of _all_pages_data after extraction
                    self._print_all_pages_data_summary("extract_specific_region")

                # Make sure to update the page_regions and page_column_lines dictionaries
                if self.multi_page_mode:
                    # Save the current regions and column lines to the page dictionaries
                    self.page_regions[self.current_page_index] = self.regions.copy()
                    self.page_column_lines[self.current_page_index] = self.column_lines.copy()

                    # Update the regions dictionary to match page_regions for the current page
                    # This ensures that the labels are properly preserved when switching pages
                    self.regions = self.page_regions[self.current_page_index]

                    print(f"[DEBUG] Updated page_regions and page_column_lines for page {self.current_page_index + 1}")

                # Update the cached extraction data with the current data
                # We need to preserve data from all pages, so we'll update only the specific section
                if self.multi_page_mode and hasattr(self, '_cached_extraction_data') and self._cached_extraction_data:
                    # Get the existing cached data
                    cached_data = copy.deepcopy(self._cached_extraction_data)

                    # CRITICAL: Use cached multi-page extraction results if available
                    # This ensures that all regions from all pages are preserved
                    from pdf_extraction_utils import get_multipage_extraction
                    multipage_cached_data = get_multipage_extraction(self.pdf_path)
                    if multipage_cached_data:
                        print(f"[DEBUG] Using cached multi-page extraction results in extract_specific_region")
                        # Make a deep copy to avoid modifying the original
                        cached_data = copy.deepcopy(multipage_cached_data)

                    # Now update only the specific section for the current page
                    if section_type == 'header':
                        # If header is a list, we need to find the right index to update
                        if isinstance(cached_data.get('header'), list) and isinstance(current_data.get('header'), list):
                            # Find the index of the region to update based on region label and page number
                            updated = False
                            for i, df in enumerate(cached_data['header']):
                                if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns and 'page_number' in df.columns:
                                    # Check if this DataFrame is for the current page and region
                                    first_label = df['region_label'].iloc[0] if len(df) > 0 else ""
                                    page_num = df['page_number'].iloc[0] if len(df) > 0 else 0

                                    # Extract region type from label (e.g., H1 from H1_R1_P1)
                                    region_type_match = re.match(r'^([HIS]\d+)_', first_label)
                                    region_type_str = region_type_match.group(1) if region_type_match else ""

                                    # Check if this is the region we're updating
                                    if page_num == self.current_page_index + 1 and region_type_str == region_label:
                                        # Update this DataFrame
                                        if region_index < len(current_data['header']):
                                            cached_data['header'][i] = current_data['header'][region_index]
                                            updated = True
                                            print(f"[DEBUG] Updated header region {region_label} on page {page_num} in cached data")
                                            break

                            # If we didn't find a matching region to update, add the new region
                            if not updated and region_index < len(current_data['header']):
                                cached_data['header'].append(current_data['header'][region_index])
                                print(f"[DEBUG] Added new header region {region_label} on page {self.current_page_index + 1} to cached data")
                        else:
                            # If header is not a list, just use the current data
                            cached_data['header'] = current_data['header']
                            print(f"[DEBUG] Replaced header in cached data with current data")

                    # Do the same for items and summary if needed
                    if section_type == 'items':
                        # Handle items section similarly
                        if isinstance(cached_data.get('items'), list) and isinstance(current_data.get('items'), list):
                            # Find the index of the region to update based on region label and page number
                            updated = False
                            for i, df in enumerate(cached_data['items']):
                                if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns and 'page_number' in df.columns:
                                    # Check if this DataFrame is for the current page and region
                                    first_label = df['region_label'].iloc[0] if len(df) > 0 else ""
                                    page_num = df['page_number'].iloc[0] if len(df) > 0 else 0

                                    # Extract region type from label (e.g., I1 from I1_R1_P1)
                                    region_type_match = re.match(r'^([HIS]\d+)_', first_label)
                                    region_type_str = region_type_match.group(1) if region_type_match else ""

                                    # Check if this is the region we're updating
                                    if page_num == self.current_page_index + 1 and region_type_str == region_label:
                                        # Update this DataFrame
                                        if region_index < len(current_data['items']):
                                            cached_data['items'][i] = current_data['items'][region_index]
                                            updated = True
                                            print(f"[DEBUG] Updated items region {region_label} on page {page_num} in cached data")
                                            break

                            # If we didn't find a matching region to update, add the new region
                            if not updated and region_index < len(current_data['items']):
                                cached_data['items'].append(current_data['items'][region_index])
                                print(f"[DEBUG] Added new items region {region_label} on page {self.current_page_index + 1} to cached data")
                        else:
                            # If items is not a list, just use the current data
                            cached_data['items'] = current_data['items']
                            print(f"[DEBUG] Replaced items in cached data with current data")

                    if section_type == 'summary':
                        # Handle summary section similarly
                        if isinstance(cached_data.get('summary'), list) and isinstance(current_data.get('summary'), list):
                            # Find the index of the region to update based on region label and page number
                            updated = False
                            for i, df in enumerate(cached_data['summary']):
                                if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns and 'page_number' in df.columns:
                                    # Check if this DataFrame is for the current page and region
                                    first_label = df['region_label'].iloc[0] if len(df) > 0 else ""
                                    page_num = df['page_number'].iloc[0] if len(df) > 0 else 0

                                    # Extract region type from label (e.g., S1 from S1_R1_P1)
                                    region_type_match = re.match(r'^([HIS]\d+)_', first_label)
                                    region_type_str = region_type_match.group(1) if region_type_match else ""

                                    # Check if this is the region we're updating
                                    if page_num == self.current_page_index + 1 and region_type_str == region_label:
                                        # Update this DataFrame
                                        if region_index < len(current_data['summary']):
                                            cached_data['summary'][i] = current_data['summary'][region_index]
                                            updated = True
                                            print(f"[DEBUG] Updated summary region {region_label} on page {page_num} in cached data")
                                            break

                            # If we didn't find a matching region to update, add the new region
                            if not updated and region_index < len(current_data['summary']):
                                cached_data['summary'].append(current_data['summary'][region_index])
                                print(f"[DEBUG] Added new summary region {region_label} on page {self.current_page_index + 1} to cached data")
                        else:
                            # If summary is not a list, just use the current data
                            cached_data['summary'] = current_data['summary']
                            print(f"[DEBUG] Replaced summary in cached data with current data")

                    # CRITICAL: Ensure the original region labels are preserved exactly as they are
                    for section in ['header', 'items', 'summary']:
                        if section in cached_data:
                            if isinstance(cached_data[section], list):
                                for i, df in enumerate(cached_data[section]):
                                    if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                                        # Make a copy of the original labels to ensure they're not modified
                                        original_labels = df['region_label'].copy()
                                        # Ensure the region labels are preserved exactly as they are
                                        df['region_label'] = original_labels

                                        # Get page number and region labels for debugging
                                        page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else "unknown"
                                        region_labels = df['region_label'].tolist() if 'region_label' in df.columns else []
                                        print(f"[DEBUG] Verified preserved region labels for {section}[{i}] from page {page_num}: {region_labels}")

                    # Update the cached data
                    self._cached_extraction_data = cached_data

                    # Store in the multi-page cache for future use
                    from pdf_extraction_utils import store_multipage_extraction
                    store_multipage_extraction(self.pdf_path, cached_data)
                    print(f"[DEBUG] Updated multi-page cache for PDF: {self.pdf_path}")

                    print(f"[DEBUG] Updated cached extraction data after specific region extraction while preserving data from all pages")
                else:
                    # In single-page mode, just use the current data
                    self._cached_extraction_data = current_data
                    print(f"[DEBUG] Updated cached extraction data after specific region extraction")

                # Update the extraction state to avoid duplicate extractions
                # In multipage mode, we need to force extraction when switching pages
                if self.multi_page_mode:
                    # Don't update _last_extraction_state in multipage mode to ensure extraction happens when switching pages
                    print(f"[DEBUG] In multipage mode, not updating _last_extraction_state to ensure extraction on page changes")
                    self._last_extraction_state = None
                else:
                    # In single-page mode, update _last_extraction_state as usual
                    self._last_extraction_state = self._get_extraction_state()

                # Update the JSON tree with the extracted data
                print(f"[DEBUG] Updating JSON tree with new data after specific region extraction")

                # Clear the tree first to ensure a clean update
                if hasattr(self, 'json_tree'):
                    self.json_tree.clear()

                # Update with current page data, even in multi-page mode
                # This ensures we only show the extraction results for the current page
                # when a specific region is extracted
                print(f"[DEBUG] Updating JSON tree with current page data only for specific region extraction")

                # Set flag to indicate we're in a specific region extraction context
                self._in_specific_region_extraction = True

                # Set flag to skip extraction update for other pages
                self._skip_extraction_update = True

                # Update the JSON tree with the current page data
                self.update_json_tree(current_data)

                # Reset the flags
                self._in_specific_region_extraction = False
                self._skip_extraction_update = False

                # Store the extracted data in the multi-page cache
                # This ensures that when we switch pages, the extraction results are preserved
                if self.multi_page_mode:
                    # Get the page number from the extracted data if available
                    page_index = self.current_page_index
                    if section_type in current_data and isinstance(current_data[section_type], pd.DataFrame) and 'page_number' in current_data[section_type].columns:
                        # Use the page number from the data (1-based) to determine the page index (0-based)
                        page_num = current_data[section_type]['page_number'].iloc[0]
                        page_index = page_num - 1
                        print(f"[DEBUG] Using page number {page_num} from data to determine page index {page_index}")

                    # Get the current multi-page extraction data
                    from pdf_extraction_utils import get_multipage_extraction, store_multipage_extraction, clear_extraction_cache_for_pdf

                    # IMPORTANT: Do NOT clear the multi-page cache for this PDF
                    # This is critical to prevent regions with the same label on different pages from being replaced
                    # Instead, get the existing multi-page data from the cache
                    cached_multipage_data = get_multipage_extraction(self.pdf_path)

                    if cached_multipage_data:
                        print(f"[DEBUG] Using cached multi-page data instead of clearing cache and re-extracting")
                        multipage_data = cached_multipage_data
                    else:
                        # Only if there's no cached data, extract it
                        print(f"[DEBUG] No cached multi-page data found, extracting from all pages")
                        multipage_data = self.extract_multi_page_invoice()

                    # Update the specific section in the multi-page data
                    new_data = current_data.get(section_type)

                    # Handle different data types
                    if isinstance(new_data, list) and new_data:
                        # For list of DataFrames, we need to update the specific region
                        if not isinstance(multipage_data.get(section_type), list):
                            multipage_data[section_type] = []

                        # Find the matching region in the multi-page data
                        updated = False
                        if region_index < len(new_data) and new_data[region_index] is not None:
                            # Get the region label from the new data
                            if isinstance(new_data[region_index], pd.DataFrame) and 'region_label' in new_data[region_index].columns:
                                # Extract both the base region name AND the page number from the region label
                                # This is critical to prevent regions with the same label on different pages from being replaced
                                new_region_label = self._get_base_region_name(new_data[region_index]['region_label'].iloc[0])

                                # Print debug info about the region label
                                print(f"[DEBUG] Processing region with label: {new_data[region_index]['region_label'].iloc[0]}")
                                print(f"[DEBUG] Extracted base region name with page: '{new_region_label}' from label: '{new_data[region_index]['region_label'].iloc[0]}'")

                                # CRITICAL: Ensure page number is set correctly
                                if 'page_number' not in new_data[region_index].columns:
                                    new_data[region_index]['page_number'] = page_index + 1
                                    print(f"[DEBUG] Added missing page_number column with value {page_index + 1}")

                                # Look for matching region in multi-page data
                                for i, df in enumerate(multipage_data.get(section_type, [])):
                                    if isinstance(df, pd.DataFrame) and 'region_label' in df.columns and 'page_number' in df.columns:
                                        existing_region_label = self._get_base_region_name(df['region_label'].iloc[0])
                                        existing_page = df['page_number'].iloc[0]

                                        print(f"[DEBUG] Comparing new region '{new_region_label}' on page {page_index + 1} with existing region '{existing_region_label}' on page {existing_page}")

                                        # Match both region label AND page number
                                        # This is critical to prevent regions with the same label on different pages from being replaced
                                        if existing_region_label == new_region_label and existing_page == page_index + 1:
                                            # Found matching region on the same page, update it
                                            multipage_data[section_type][i] = new_data[region_index]
                                            updated = True
                                            print(f"[DEBUG] Updated {section_type}[{i}] in multi-page cache for page {page_index + 1}")
                                            break

                            # If no matching region found, add the new data
                            if not updated:
                                multipage_data[section_type].append(new_data[region_index])
                                print(f"[DEBUG] Added new {section_type} region to multi-page cache for page {page_index + 1}")

                    elif isinstance(new_data, pd.DataFrame) and not new_data.empty:
                        # For single DataFrame, we need to update or add it
                        if 'region_label' in new_data.columns:
                            new_region_label = self._get_base_region_name(new_data['region_label'].iloc[0])

                            # CRITICAL: Ensure page number is set correctly
                            if 'page_number' not in new_data.columns:
                                new_data['page_number'] = page_index + 1

                            # Initialize section if needed
                            if not isinstance(multipage_data.get(section_type), list):
                                multipage_data[section_type] = []

                            # Look for matching region in multi-page data
                            updated = False
                            for i, df in enumerate(multipage_data.get(section_type, [])):
                                if isinstance(df, pd.DataFrame) and 'region_label' in df.columns and 'page_number' in df.columns:
                                    existing_region_label = self._get_base_region_name(df['region_label'].iloc[0])
                                    existing_page = df['page_number'].iloc[0]

                                    # Match both region label AND page number
                                    if existing_region_label == new_region_label and existing_page == page_index + 1:
                                        # Found matching region on the same page, update it
                                        multipage_data[section_type][i] = new_data
                                        updated = True
                                        print(f"[DEBUG] Updated {section_type}[{i}] in multi-page cache for page {page_index + 1}")
                                        break

                            # If no matching region found, add the new data
                            if not updated:
                                multipage_data[section_type].append(new_data)
                                print(f"[DEBUG] Added new {section_type} DataFrame to multi-page cache for page {page_index + 1}")
                        else:
                            # No region label, just add or replace
                            if not isinstance(multipage_data.get(section_type), list):
                                multipage_data[section_type] = []
                            multipage_data[section_type].append(new_data)
                            print(f"[DEBUG] Added {section_type} DataFrame to multi-page cache for page {page_index + 1}")

                    # Store the updated multi-page data
                    store_multipage_extraction(self.pdf_path, multipage_data)
                    print(f"[DEBUG] Updated multi-page cache for PDF: {self.pdf_path}")

                    # Print a summary of the updated multi-page data to verify all regions are preserved
                    for section in ['header', 'items', 'summary']:
                        if section in multipage_data and isinstance(multipage_data[section], list):
                            print(f"[DEBUG] {section} section has {len(multipage_data[section])} DataFrames")
                            for i, df in enumerate(multipage_data[section]):
                                if isinstance(df, pd.DataFrame) and not df.empty:
                                    if 'region_label' in df.columns and 'page_number' in df.columns:
                                        region_labels = df['region_label'].unique().tolist()
                                        page_nums = df['page_number'].unique().tolist()
                                        print(f"[DEBUG]   {section}[{i}] has region labels {region_labels} from page(s) {page_nums}")

                # Force the UI to update and ensure the tree is properly refreshed
                QApplication.processEvents()

                # Ensure the tree has focus for keyboard shortcuts
                if hasattr(self, 'json_tree'):
                    self.json_tree.setFocus()
            else:
                print(f"[DEBUG] No data extracted for {section_type} region {region_index}")

        except Exception as e:
            print(f"Error extracting specific region: {str(e)}")
            import traceback
            traceback.print_exc()

    def multipage_extract_page_data(self, page_index):
        """Extract data from a specific page with enhanced multi-page support

        This method extends extract_page_data with specific multi-page enhancements
        to ensure proper page tracking and region labeling.

        This implementation uses the _EXTRACTION_CACHE mechanism from pdf_extraction_utils
        instead of the redundant _all_pages_data structure.

        Args:
            page_index (int): The index of the PDF page (0-based)

        Returns:
            dict: Dictionary containing 'header', 'items', and 'summary' data for the page
        """
        # Save current page index
        current_page = self.current_page_index

        # Switch to the requested page if needed
        if current_page != page_index:
            self.current_page_index = page_index
            self.display_current_page()

        # First, extract the data using the standard method
        header_df, items_df, summary_df = self.extract_page_data(page_index)

        # Restore original page if needed
        if current_page != page_index:
            self.current_page_index = current_page
            self.display_current_page()

        # Initialize result dictionary with empty lists for all sections to ensure consistency
        result = {
            'header': [] if header_df is None else header_df,
            'items': [] if items_df is None else items_df,
            'summary': [] if summary_df is None else summary_df
        }

        # Add page number and region labels to each DataFrame
        page_num = page_index + 1  # 1-based page number

        # Process header DataFrame
        if isinstance(result['header'], list):
            # Handle list of DataFrames
            for i, df in enumerate(result['header']):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Ensure page number columns exist
                    if 'page_number' not in df.columns:
                        df['page_number'] = page_num
                    if '_page_number' not in df.columns:
                        df['_page_number'] = page_num

                    # Add region_label column if it doesn't exist
                    if 'region_label' not in df.columns:
                        # Use the actual region label from the region dictionary
                        region_label = f"H{i+1}"  # Default
                        if 'header' in self.page_regions.get(page_index, {}) and i < len(self.page_regions[page_index]['header']):
                            # Get the label from the corresponding header region
                            region_item = self.page_regions[page_index]['header'][i]
                            if isinstance(region_item, dict) and 'label' in region_item:
                                region_label = region_item['label']

                        # IMPORTANT: Preserve the original region label (H1, H2, etc.) exactly as it is
                        df['region_label'] = [f"{region_label}_R{j+1}_P{page_num}"
                                           for j in range(len(df))]
                    else:
                        # DO NOT modify region labels - preserve them exactly as they are
                        # Make a copy of the original labels to ensure they're not modified
                        original_labels = df['region_label'].copy()

                        # Ensure the region labels are preserved exactly as they are
                        df['region_label'] = original_labels

        # Process items DataFrame
        if isinstance(result['items'], list):
            # Handle list of DataFrames
            for i, df in enumerate(result['items']):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Ensure page number columns exist
                    if 'page_number' not in df.columns:
                        df['page_number'] = page_num
                    if '_page_number' not in df.columns:
                        df['_page_number'] = page_num

                    # Add region_label column if it doesn't exist
                    if 'region_label' not in df.columns:
                        # Use the actual region label from the region dictionary
                        region_label = f"I{i+1}"  # Default
                        if 'items' in self.page_regions.get(page_index, {}) and i < len(self.page_regions[page_index]['items']):
                            # Get the label from the corresponding items region
                            region_item = self.page_regions[page_index]['items'][i]
                            if isinstance(region_item, dict) and 'label' in region_item:
                                region_label = region_item['label']

                        # IMPORTANT: Preserve the original region label (I1, I2, etc.) exactly as it is
                        df['region_label'] = [f"{region_label}_R{j+1}_P{page_num}"
                                           for j in range(len(df))]
                    else:
                        # DO NOT modify region labels - preserve them exactly as they are
                        # Make a copy of the original labels to ensure they're not modified
                        original_labels = df['region_label'].copy()

                        # Ensure the region labels are preserved exactly as they are
                        df['region_label'] = original_labels

        # Process summary DataFrame
        if isinstance(result['summary'], list):
            # Handle list of DataFrames
            for i, df in enumerate(result['summary']):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Ensure page number columns exist
                    if 'page_number' not in df.columns:
                        df['page_number'] = page_num
                    if '_page_number' not in df.columns:
                        df['_page_number'] = page_num

                    # Add region_label column if it doesn't exist
                    if 'region_label' not in df.columns:
                        # Use the actual region label from the region dictionary
                        region_label = f"S{i+1}"  # Default
                        if 'summary' in self.page_regions.get(page_index, {}) and i < len(self.page_regions[page_index]['summary']):
                            # Get the label from the corresponding summary region
                            region_item = self.page_regions[page_index]['summary'][i]
                            if isinstance(region_item, dict) and 'label' in region_item:
                                region_label = region_item['label']

                        # IMPORTANT: Preserve the original region label (S1, S2, etc.) exactly as it is
                        df['region_label'] = [f"{region_label}_R{j+1}_P{page_num}"
                                           for j in range(len(df))]
                    else:
                        # DO NOT modify region labels - preserve them exactly as they are
                        # Make a copy of the original labels to ensure they're not modified
                        original_labels = df['region_label'].copy()

                        # Ensure the region labels are preserved exactly as they are
                        df['region_label'] = original_labels

        # Return the data as a dictionary with consistent structure
        return result

    def extract_with_unified_engine(self, use_cache: bool = True) -> Dict[str, Any]:
        """
        Extract invoice tables using the unified extraction engine.

        This method uses the new common extraction engine with region adapters
        to provide identical extraction logic as bulk_processor.py while
        maintaining the UI-based region input.

        Args:
            use_cache (bool): Whether to use extraction cache

        Returns:
            dict: Extraction results with standardized format
        """
        if not UNIFIED_EXTRACTION_AVAILABLE:
            print("[WARNING] Unified extraction not available, falling back to legacy method")
            return self.extract_multi_page_invoice()

        try:
            print(f"[DEBUG] Starting unified extraction for {self.pdf_path}")

            # Use the unified extraction adapter
            results = extract_from_user_regions(
                pdf_path=self.pdf_path,
                regions=self.regions,
                column_lines=self.column_lines,
                extraction_params=getattr(self, 'extraction_params', {}),
                multi_page_mode=self.multi_page_mode,
                page_regions=getattr(self, 'page_regions', {}),
                page_column_lines=getattr(self, 'page_column_lines', {}),
                use_cache=use_cache,
                processor_instance=self
            )

            if results:
                print(f"[DEBUG] Unified extraction completed successfully")

                # Convert to legacy format for compatibility
                legacy_results = self._convert_unified_to_legacy_format(results)

                # Update cached data
                self._cached_extraction_data = legacy_results

                # Store in multi-page cache
                if self.multi_page_mode:
                    from pdf_extraction_utils import store_multipage_extraction
                    store_multipage_extraction(self.pdf_path, legacy_results)

                return legacy_results
            else:
                print("[WARNING] Unified extraction returned no results")
                return self._create_empty_extraction_result()

        except Exception as e:
            print(f"[ERROR] Unified extraction failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return self._create_empty_extraction_result()

    def _convert_unified_to_legacy_format(self, unified_results: Dict[str, Any]) -> Dict[str, Any]:
        """
        Convert unified extraction results to legacy format for compatibility.

        Args:
            unified_results (dict): Results from unified extraction engine

        Returns:
            dict: Results in legacy format
        """
        try:
            legacy_results = {
                'header': unified_results.get('header_tables', []),
                'items': unified_results.get('items_tables', []),
                'summary': unified_results.get('summary_tables', []),
                'metadata': unified_results.get('metadata', {})
            }

            # Ensure lists are not empty lists when no data
            for section in ['header', 'items', 'summary']:
                if not legacy_results[section]:
                    legacy_results[section] = None
                elif len(legacy_results[section]) == 1:
                    # Convert single-item lists to single items for backward compatibility
                    legacy_results[section] = legacy_results[section][0]

            return legacy_results

        except Exception as e:
            print(f"[ERROR] Failed to convert unified results to legacy format: {str(e)}")
            return self._create_empty_extraction_result()

    def _create_empty_extraction_result(self) -> Dict[str, Any]:
        """Create empty extraction result structure."""
        return {
            'header': None,
            'items': None,
            'summary': None,
            'metadata': {
                'filename': os.path.basename(self.pdf_path) if self.pdf_path else '',
                'page_count': len(self.pdf_document) if self.pdf_document else 0,
                'template_type': 'multi' if self.multi_page_mode else 'single'
            }
        }

    def initialize_invoice2data_template(self):
        """Initialize the invoice2data template with default values and populate the fields tables"""
        try:
            # Create a default template if none exists using factory
            if not hasattr(self, 'invoice2data_template') or self.invoice2data_template is None:
                # Use factory to create default template with enhanced structure
                base_template = TemplateFactory.create_default_invoice_template("Company Name", "USD")

                # Enhance with invoice2data specific fields
                self.invoice2data_template = {
                    "issuer": base_template["issuer"],
                    "fields": {
                        "invoice_number": {
                            "parser": "regex",
                            "regex": "Invoice Number:\\s*([\\w-]+)",
                            "type": "string"
                        },
                        "date": {
                            "parser": "regex",
                            "regex": "Date:\\s*(\\d{4}-\\d{2}-\\d{2})",
                            "type": "date",
                            "formats": ["%Y-%m-%d"]
                        },
                        "amount": {
                            "parser": "regex",
                            "regex": "Amount:\\s*\\$(\\d+\\.\\d{2})",
                            "type": "float"
                        }
                    },
                    "keywords": base_template["keywords"],
                    "options": base_template["options"]
                }
                print(f"[DEBUG] Created default invoice2data template using factory")

            # Directly populate the header fields table
            if hasattr(self, 'header_fields_table') and 'fields' in self.invoice2data_template:
                fields = self.invoice2data_template['fields']

                # Clear existing fields
                while self.header_fields_table.rowCount() > 0:
                    self.header_fields_table.removeRow(0)

                # Add fields from template
                for field_name, field_data in fields.items():
                    row = self.header_fields_table.rowCount()
                    self.header_fields_table.insertRow(row)

                    # Set field name
                    self.header_fields_table.setItem(row, 0, QTableWidgetItem(field_name))

                    # Set regex pattern
                    if isinstance(field_data, dict) and 'regex' in field_data:
                        regex_pattern = field_data['regex']
                    else:
                        # Simple field format (just a regex pattern)
                        regex_pattern = str(field_data)
                    self.header_fields_table.setItem(row, 1, QTableWidgetItem(regex_pattern))

                    # Set field type
                    type_combo = QComboBox()
                    type_combo.addItems(["string", "date", "float", "int"])
                    if isinstance(field_data, dict) and 'type' in field_data:
                        field_type = field_data['type']
                        if field_type in ["string", "date", "float", "int"]:
                            type_combo.setCurrentText(field_type)
                    self.header_fields_table.setCellWidget(row, 2, type_combo)

                print(f"[DEBUG] Directly populated header fields table with {self.header_fields_table.rowCount()} fields")

            # Also populate the summary fields table
            if hasattr(self, 'summary_fields_table') and 'fields' in self.invoice2data_template:
                # Use the populate_form_from_template method to populate the summary fields table
                self.populate_form_from_template()
                print(f"[DEBUG] Populated form from template")

        except Exception as e:
            print(f"[ERROR] Failed to initialize invoice2data template: {str(e)}")
            import traceback
            traceback.print_exc()

    def _print_all_pages_data_summary(self, caller_name=""):
        """Print a summary of the extraction cache for debugging purposes.

        Args:
            caller_name (str): Name of the calling method for logging purposes
        """
        # Get cache statistics
        from pdf_extraction_utils import get_extraction_cache_stats
        stats = get_extraction_cache_stats()
        print(f"[DEBUG] Extraction cache stats from {caller_name}: {stats}")

    def _get_base_region_name(self, label):
        """Extract the base region name from a label.

        For example, from "H1_R1_P1" it extracts "H1_P1" to preserve page information.

        In multi-page PDF extraction mode, we preserve region data from all pages without duplicate checking,
        as the same region can appear on different pages. This method ensures that we correctly identify
        the base region name WITH the page number to prevent regions from different pages being confused.

        CRITICAL: This method is essential to prevent regions with the same label on different pages
        from being replaced. It ensures that we preserve both the region type (H1, H2) AND the page
        number (P1, P2) when comparing regions.

        Args:
            label (str): The region label

        Returns:
            str: The base region name with page number
        """
        # Extract the base region name (without row number but WITH page number)
        if not label:
            return ""

        # The format is typically like "H1_R1_P1" or "H2_R1_P1" where H1 or H2 is the base region name
        # We need to preserve both the region number (H1, H2) AND the page number (P1, P2)
        parts = label.split('_')
        if len(parts) >= 3 and parts[0] and parts[2]:
            # CRITICAL: Return format: "H1_P1" to preserve both region type and page number
            # This is essential to prevent regions with the same label on different pages from being replaced
            base_with_page = f"{parts[0]}_{parts[2]}"
            print(f"[DEBUG] Extracted base region name with page: '{base_with_page}' from label: '{label}'")
            return base_with_page
        elif len(parts) > 0:
            # Fallback to just the region type if no page info
            base_name = parts[0]
            print(f"[DEBUG] Extracted base region name without page: '{base_name}' from label: '{label}'")
            return base_name

        print(f"[DEBUG] Could not extract base region name from label: '{label}', returning as is")
        return label

    def extract_multi_page_invoice(self):
        """Extract and combine data from all pages of a multi-page invoice.

        This method preserves each page's data separately and combines them by section.
        It ensures that region labels are preserved exactly as they are, and data from
        different pages with the same region label is not replaced.

        This implementation uses the _EXTRACTION_CACHE mechanism from pdf_extraction_utils
        instead of the redundant _all_pages_data structure.

        Returns:
            dict: Dictionary containing combined data from all pages
        """
        if not self.pdf_document:
            return {'header': None, 'items': None, 'summary': None}

        # Use cached data if available
        if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data:
            # Make a deep copy to avoid modifying the original
            cached_copy = copy.deepcopy(self._cached_extraction_data)

            # CRITICAL: Ensure the original region labels are preserved exactly as they are
            # This is essential to prevent regions with the same label on different pages from being replaced
            for section in ['header', 'items', 'summary']:
                if section in cached_copy:
                    if isinstance(cached_copy[section], list):
                        for i, df in enumerate(cached_copy[section]):
                            if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                                # Make a copy of the original labels to ensure they're not modified
                                original_labels = df['region_label'].copy()
                                # Ensure the region labels are preserved exactly as they are
                                df['region_label'] = original_labels

                                # Get page number and region labels for debugging
                                page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else "unknown"
                                region_labels = df['region_label'].tolist() if 'region_label' in df.columns else []
                                print(f"[DEBUG] Verified preserved region labels for {section}[{i}] from page {page_num}: {region_labels}")

            return cached_copy

        # Check if we have cached multi-page extraction results
        from pdf_extraction_utils import get_multipage_extraction, store_multipage_extraction
        cached_data = get_multipage_extraction(self.pdf_path)
        if cached_data:
            # Make a deep copy to avoid modifying the original
            cached_copy = copy.deepcopy(cached_data)

            # CRITICAL: Ensure the original region labels are preserved exactly as they are
            # This is essential to prevent regions with the same label on different pages from being replaced
            for section in ['header', 'items', 'summary']:
                if section in cached_copy:
                    if isinstance(cached_copy[section], list):
                        for i, df in enumerate(cached_copy[section]):
                            if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                                # Make a copy of the original labels to ensure they're not modified
                                original_labels = df['region_label'].copy()
                                # Ensure the region labels are preserved exactly as they are
                                df['region_label'] = original_labels

                                # Get page number and region labels for debugging
                                page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else "unknown"
                                region_labels = df['region_label'].tolist() if 'region_label' in df.columns else []
                                print(f"[DEBUG] Verified preserved region labels for {section}[{i}] from page {page_num}: {region_labels}")

            # Store in instance cache for faster access next time
            self._cached_extraction_data = copy.deepcopy(cached_copy)
            return cached_copy

        # Extract data for all pages if needed
        all_dataframes = []  # Will store tuples of (page_idx, section, region_type, df)

        # Extract data for each page
        for page_idx in range(len(self.pdf_document)):
            # Extract data for this page
            page_data = self.multipage_extract_page_data(page_idx)

            # Process each section in the page data
            for section in ['header', 'items', 'summary']:
                section_data = page_data.get(section)

                # Handle list of DataFrames
                if isinstance(section_data, list):
                    for df_idx, df in enumerate(section_data):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            # Create a copy to avoid modifying the original
                            df_copy = df.copy()

                            # Ensure page_number column exists
                            if 'page_number' not in df_copy.columns:
                                df_copy['page_number'] = page_idx + 1

                            # Get region type from region_label if available
                            region_type = f"{section[0].upper()}{df_idx+1}"  # Default
                            if 'region_label' in df_copy.columns and len(df_copy) > 0:
                                first_label = df_copy['region_label'].iloc[0]
                                if isinstance(first_label, str) and len(first_label) > 1:
                                    # Extract the region type (H1, H2, etc.)
                                    region_match = re.match(r'^([HIS]\d+)_', first_label)
                                    if region_match:
                                        region_type = region_match.group(1)

                            # Add to collection with metadata
                            # Store the original page index to ensure proper sorting
                            # CRITICAL: Include page_idx in the tuple to ensure regions with same label on different pages are preserved
                            all_dataframes.append((page_idx, section, region_type, df_copy))

                # Handle single DataFrame
                elif isinstance(section_data, pd.DataFrame) and not section_data.empty:
                    # Create a copy to avoid modifying the original
                    df_copy = section_data.copy()

                    # Ensure page_number column exists
                    if 'page_number' not in df_copy.columns:
                        df_copy['page_number'] = page_idx + 1

                    # Get region type from region_label if available
                    region_type = f"{section[0].upper()}1"  # Default
                    if 'region_label' in df_copy.columns and len(df_copy) > 0:
                        first_label = df_copy['region_label'].iloc[0]
                        if isinstance(first_label, str) and len(first_label) > 1:
                            # Extract the region type (H1, H2, etc.)
                            region_match = re.match(r'^([HIS]\d+)_', first_label)
                            if region_match:
                                region_type = region_match.group(1)

                    # Add to collection with metadata
                    all_dataframes.append((page_idx, section, region_type, df_copy))

        # Group DataFrames by section, preserving page-specific data
        header_dfs = []
        item_dfs = []
        summary_dfs = []

        # CRITICAL CHANGE: Sort by section first, then by page, then by region type
        # This ensures that data from different pages is kept separate even if they have the same region type
        # This is essential to prevent regions with the same label on different pages from being grouped incorrectly
        #
        # The tuple structure is: (page_idx, section, region_type, df)
        # So we sort by:
        # 1. section (x[1]) - to group by header/items/summary
        # 2. page_idx (x[0]) - to keep pages separate
        # 3. region_type (x[2]) - to order regions within a page
        #
        # This sorting ensures that regions with the same label (e.g., H1) on different pages
        # are kept separate and not replaced by each other
        all_dataframes.sort(key=lambda x: (x[1], x[0], x[2]))

        # Debug print to show how regions are being sorted
        print(f"[DEBUG] Sorted dataframes: {[(x[0]+1, x[1], x[2]) for x in all_dataframes]}")

        # Process each DataFrame and add to the appropriate section
        for page_idx, section, region_type, df in all_dataframes:
            if section == 'header':
                print(f"[DEBUG] Adding header region from page {page_idx+1} with region type {region_type}")
                header_dfs.append(df)
            elif section == 'items':
                print(f"[DEBUG] Adding items region from page {page_idx+1} with region type {region_type}")
                item_dfs.append(df)
            elif section == 'summary':
                print(f"[DEBUG] Adding summary region from page {page_idx+1} with region type {region_type}")
                summary_dfs.append(df)

        # Create result structure with all DataFrames preserved separately
        result = {
            'header': header_dfs if header_dfs else None,
            'items': item_dfs if item_dfs else None,
            'summary': summary_dfs if summary_dfs else None
        }

        # Get existing metadata if available to preserve creation_date
        existing_metadata = None
        if hasattr(self, '_cached_extraction_data') and self._cached_extraction_data is not None and 'metadata' in self._cached_extraction_data:
            existing_metadata = self._cached_extraction_data['metadata']

        # Add metadata
        if self.pdf_path:
            if existing_metadata and 'creation_date' in existing_metadata:
                # Preserve the original creation_date to prevent region label changes
                result['metadata'] = {
                    'filename': os.path.basename(self.pdf_path),
                    'page_count': len(self.pdf_document),
                    'template_type': 'multi',
                    'creation_date': existing_metadata['creation_date']
                }
            else:
                # Only set a new creation_date if one doesn't exist
                result['metadata'] = {
                    'filename': os.path.basename(self.pdf_path),
                    'page_count': len(self.pdf_document),
                    'template_type': 'multi',
                    'creation_date': datetime.datetime.now().isoformat()
                }

        # CRITICAL: Ensure the original region labels are preserved exactly as they are
        for section in ['header', 'items', 'summary']:
            if section in result:
                if isinstance(result[section], list):
                    for i, df in enumerate(result[section]):
                        if isinstance(df, pd.DataFrame) and not df.empty and 'region_label' in df.columns:
                            # Make a copy of the original labels to ensure they're not modified
                            original_labels = df['region_label'].copy()
                            # Ensure the region labels are preserved exactly as they are
                            df['region_label'] = original_labels

                            # Get page number and region labels for debugging
                            page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else "unknown"
                            region_labels = df['region_label'].tolist() if 'region_label' in df.columns else []
                            print(f"[DEBUG] Preserved original region labels for {section}[{i}] from page {page_num}: {region_labels}")
                elif isinstance(result[section], pd.DataFrame) and not result[section].empty and 'region_label' in result[section].columns:
                    # Make a copy of the original labels to ensure they're not modified
                    original_labels = result[section]['region_label'].copy()
                    # Ensure the region labels are preserved exactly as they are
                    result[section]['region_label'] = original_labels

                    # Get page numbers and region labels for debugging
                    page_nums = result[section]['page_number'].unique().tolist() if 'page_number' in result[section].columns else []
                    region_labels = result[section]['region_label'].tolist() if 'region_label' in result[section].columns else []
                    print(f"[DEBUG] Preserved original region labels for {section} from pages {page_nums}: {region_labels}")

        # Cache the combined data for future use
        self._cached_extraction_data = copy.deepcopy(result)

        # Store in the multi-page cache for future use
        store_multipage_extraction(self.pdf_path, result)

        return result

    # multipage_extract_invoice method has been removed as it was deprecated

    def save_template_directly(self, name, description, use_middle_page=False, fixed_page_count=False, **additional_params):
        """Save the template directly to the database with actual parameters

        Args:
            name (str): Template name
            description (str): Template description
            use_middle_page (bool): Whether to use only first, middle, and last pages (for multi-page templates)
            fixed_page_count (bool): Whether to enforce fixed page count (for multi-page templates)
            **additional_params: Additional parameters to store in the config
        """
        try:
            from database import InvoiceDatabase
            from PySide6.QtWidgets import QMessageBox
            import os

            print("\nStarting template save process...")

            # Validate input parameters
            if not name or not isinstance(name, str):
                raise ValueError("Template name must be a non-empty string")
            if not isinstance(description, str):
                description = str(description)

            # Open database connection
            db = InvoiceDatabase()

            # Handle multi-page templates
            if self.multi_page_mode and len(self.pdf_document) > 1:
                print(f"[DEBUG] Saving multi-page template with {len(self.pdf_document)} pages")

                # Initialize config dictionary
                config = {
                    'use_middle_page': use_middle_page,
                    'fixed_page_count': fixed_page_count,
                    'total_pages': len(self.pdf_document),
                    'page_indices': list(range(len(self.pdf_document)))
                }

                # Add extraction parameters to config
                for key, value in self.extraction_params.items():
                    if key not in config:  # Don't overwrite existing keys
                        config[key] = value

                # Add additional parameters to config
                for key, value in additional_params.items():
                    config[key] = value

                # Create page_configs as a list to match multi_page_section_viewer.py format
                page_configs = []
                print(f"[DEBUG] Creating page_configs for {len(self.pdf_document)} pages")
                print(f"[DEBUG] page_regions keys: {list(self.page_regions.keys())}")
                print(f"[DEBUG] page_column_lines keys: {list(self.page_column_lines.keys())}")

                # Make sure all pages have regions and column lines
                # This is important for the "Apply to All" button to work correctly
                for page_idx in range(len(self.pdf_document)):
                    # If a page doesn't have regions or column lines, copy from the current page
                    if page_idx not in self.page_regions:
                        print(f"[DEBUG] Page {page_idx + 1} doesn't have regions, copying from current page")
                        if self.current_page_index in self.page_regions:
                            self.page_regions[page_idx] = self.page_regions[self.current_page_index].copy()
                        elif hasattr(self, 'regions') and self.regions:
                            self.page_regions[page_idx] = self.regions.copy()
                        else:
                            # Initialize with empty regions
                            self.page_regions[page_idx] = {'header': [], 'items': [], 'summary': []}

                    if page_idx not in self.page_column_lines:
                        print(f"[DEBUG] Page {page_idx + 1} doesn't have column lines, copying from current page")
                        if self.current_page_index in self.page_column_lines:
                            self.page_column_lines[page_idx] = self.page_column_lines[self.current_page_index].copy()
                        elif hasattr(self, 'column_lines') and self.column_lines:
                            self.page_column_lines[page_idx] = self.column_lines.copy()
                        else:
                            # Initialize with empty column lines
                            self.page_column_lines[page_idx] = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}

                # Now create the page_configs with all pages
                for page_idx in range(len(self.pdf_document)):
                    # Get scale factors for this page
                    scale_factors = get_scale_factors(self.pdf_path, page_idx)

                    # Get rendered dimensions
                    rendered_width = self.pdf_document[page_idx].rect.width
                    rendered_height = self.pdf_document[page_idx].rect.height

                    # Create page config with original coordinates
                    page_config = {
                        'page_index': page_idx,
                        'scale_factors': {
                            'scale_x': scale_factors['scale_x'],
                            'scale_y': scale_factors['scale_y'],
                            'rendered_width': rendered_width,
                            'rendered_height': rendered_height,
                            'page_width': scale_factors['page_width'],
                            'page_height': scale_factors['page_height']
                        }
                    }

                    # Add original regions for this page
                    original_regions = {}
                    for section, rects in self.page_regions[page_idx].items():
                        original_regions[section] = []
                        for rect in rects:
                            # Check if rect is a QRect object or a dictionary with 'rect' key
                            if isinstance(rect, dict) and 'rect' in rect:
                                # This is a dictionary with a 'rect' key containing the QRect object
                                qrect = rect['rect']
                                original_regions[section].append({
                                    'x': qrect.x(),
                                    'y': qrect.y(),
                                    'width': qrect.width(),
                                    'height': qrect.height()
                                })
                            elif hasattr(rect, 'x') and callable(rect.x):
                                # This is a QRect object directly
                                original_regions[section].append({
                                    'x': rect.x(),
                                    'y': rect.y(),
                                    'width': rect.width(),
                                    'height': rect.height()
                                })
                            else:
                                # This is some other type of object, log it for debugging
                                print(f"[DEBUG] Unexpected rect type: {type(rect)}, skipping")
                    page_config['original_regions'] = original_regions
                    print(f"[DEBUG] Added original_regions for page {page_idx + 1}: {list(original_regions.keys())}")

                    # Add original column lines for this page
                    original_column_lines = {}
                    for section, lines in self.page_column_lines[page_idx].items():
                        section_name = section.value if hasattr(section, 'value') else section
                        original_column_lines[section_name] = []
                        for line in lines:
                            if len(line) >= 2 and isinstance(line[0], QPoint) and isinstance(line[1], QPoint):
                                line_data = [
                                    {'x': line[0].x(), 'y': line[0].y()},
                                    {'x': line[1].x(), 'y': line[1].y()}
                                ]
                                if len(line) > 2:
                                    line_data.append(line[2])  # Add region index if present
                                original_column_lines[section_name].append(line_data)
                    page_config['original_column_lines'] = original_column_lines
                    print(f"[DEBUG] Added original_column_lines for page {page_idx + 1}: {list(original_column_lines.keys())}")

                    # Add the page config to the list
                    page_configs.append(page_config)

                # Store page_configs in the instance and in the config
                self.page_configs = page_configs

                # Store page_configs in config
                config['page_configs'] = self.page_configs
                print(f"[DEBUG] Saving {len(self.page_configs)} page_configs")
                for i, page_config in enumerate(self.page_configs):
                    print(f"[DEBUG] Page config {i+1} has regions: {list(page_config.get('original_regions', {}).keys())}")

                # Convert page regions to serializable format for database as a list (not a dict)
                # This matches the format used in multi_page_section_viewer.py
                page_regions = []

                # Determine the total number of pages
                total_pages = len(self.pdf_document)

                # Initialize page_regions with empty dictionaries for each page
                for _ in range(total_pages):
                    page_regions.append({})

                # Fill in the regions for each page
                print(f"[DEBUG] page_regions keys: {list(self.page_regions.keys())}")
                for page_idx in range(total_pages):
                    if page_idx in self.page_regions:
                        # Get scale factors for this page
                        scale_factors = get_scale_factors(self.pdf_path, page_idx)
                        scale_x = scale_factors['scale_x']
                        scale_y = scale_factors['scale_y']
                        page_height = scale_factors['page_height']

                        # Create serializable regions for this page
                        serializable_regions = {}
                        for section, rects in self.page_regions[page_idx].items():
                            serializable_regions[section] = []
                            for rect in rects:
                                # Check if rect is a QRect object or a dictionary with 'rect' key
                                if isinstance(rect, dict) and 'rect' in rect:
                                    # This is a dictionary with a 'rect' key containing the QRect object
                                    qrect = rect['rect']
                                    # Convert QRect to bottom-left coordinate system
                                    x1 = qrect.x() * scale_x
                                    y1 = page_height - (qrect.y() * scale_y)  # Convert top y1 to bottom y1
                                    x2 = (qrect.x() + qrect.width()) * scale_x
                                    y2 = page_height - ((qrect.y() + qrect.height()) * scale_y)  # Convert bottom y2 to bottom y2
                                elif hasattr(rect, 'x') and callable(rect.x):
                                    # This is a QRect object directly
                                    # Convert QRect to bottom-left coordinate system
                                    x1 = rect.x() * scale_x
                                    y1 = page_height - (rect.y() * scale_y)  # Convert top y1 to bottom y1
                                    x2 = (rect.x() + rect.width()) * scale_x
                                    y2 = page_height - ((rect.y() + rect.height()) * scale_y)  # Convert bottom y2 to bottom y2
                                else:
                                    # This is some other type of object, log it for debugging
                                    print(f"[DEBUG] Unexpected rect type: {type(rect)}, skipping")
                                    continue

                                # Add to serializable regions in the format expected by the database
                                serializable_regions[section].append({
                                    'x1': x1,
                                    'y1': y1,
                                    'x2': x2,
                                    'y2': y2
                                })
                        page_regions[page_idx] = serializable_regions
                        print(f"[DEBUG] Saved regions for page {page_idx + 1}: {len(serializable_regions)}")

                # Convert page column lines to serializable format for database as a list (not a dict)
                # This matches the format used in multi_page_section_viewer.py
                page_column_lines = []

                # Initialize page_column_lines with empty dictionaries for each page
                for _ in range(total_pages):
                    page_column_lines.append({})

                # Fill in the column lines for each page
                print(f"[DEBUG] page_column_lines keys: {list(self.page_column_lines.keys())}")
                for page_idx in range(total_pages):
                    if page_idx in self.page_column_lines:
                        # Get scale factors for this page
                        scale_factors = get_scale_factors(self.pdf_path, page_idx)
                        scale_x = scale_factors['scale_x']
                        scale_y = scale_factors['scale_y']
                        page_height = scale_factors['page_height']

                        # Create serializable column lines for this page
                        serializable_column_lines = {}
                        for section, lines in self.page_column_lines[page_idx].items():
                            section_name = section.value if hasattr(section, 'value') else section
                            serializable_column_lines[section_name] = []

                            for line in lines:
                                if len(line) >= 2 and isinstance(line[0], QPoint) and isinstance(line[1], QPoint):
                                    # Convert line to serializable format with scaled coordinates
                                    # Make sure we're creating a list of QPoints for the column lines
                                    # This is the format expected by the template application code
                                    start_point = QPoint(int(line[0].x() * scale_x), int(page_height - (line[0].y() * scale_y)))
                                    end_point = QPoint(int(line[1].x() * scale_x), int(page_height - (line[1].y() * scale_y)))

                                    # Create the line data
                                    if len(line) > 2:
                                        # Include region index if present
                                        line_data = [start_point, end_point, line[2]]
                                    else:
                                        line_data = [start_point, end_point]

                                    serializable_column_lines[section_name].append(line_data)
                        page_column_lines[page_idx] = serializable_column_lines
                        print(f"[DEBUG] Saved column lines for page {page_idx + 1}: {len(serializable_column_lines)}")

                # Save template to the database
                # Get the invoice2data template JSON from the template_preview
                print(f"\n[DEBUG] Getting invoice2data template JSON for save_template_directly (multi-page)")
                json_template = None

                # Get the template directly from the template_preview as YAML
                print(f"[DEBUG] Has template_preview: {hasattr(self, 'template_preview')}")
                if hasattr(self, 'template_preview') and self.template_preview:
                    try:
                        # Get the raw text from the template preview
                        template_text = self.template_preview.toPlainText()
                        print(f"[DEBUG] Template text length: {len(template_text)}")
                        print(f"[DEBUG] Template text (first 100 chars): {template_text[:100]}...")

                        # Check if the template preview contains test results
                        if template_text.strip().startswith("TEST RESULT"):
                            # This is a test result, not a valid template
                            print(f"[DEBUG] Template preview contains test results, not a valid template")
                            # Get the last valid template that was tested
                            if hasattr(self, '_last_valid_template') and self._last_valid_template:
                                yaml_template = self._last_valid_template
                                print(f"[DEBUG] Using last valid template that was tested")
                            else:
                                # Build a new template from the form
                                yaml_template = self.build_invoice2data_template()
                                print(f"[DEBUG] No last valid template found, building from form")
                        elif template_text.strip():
                            # Try to parse as YAML first (preferred format)
                            import yaml
                            try:
                                # Parse the template to validate it's proper YAML
                                parsed_template = yaml.safe_load(template_text)
                                if parsed_template and isinstance(parsed_template, dict):
                                    # Use the parsed template directly without any modifications
                                    yaml_template = parsed_template
                                    print(f"[DEBUG] Successfully parsed template as YAML: {list(yaml_template.keys())}")
                                else:
                                    print(f"[DEBUG] Template is not a valid dictionary: {type(parsed_template)}")
                                    yaml_template = self.build_invoice2data_template()
                            except yaml.YAMLError as yaml_err:
                                # If YAML parsing fails, try JSON as fallback
                                try:
                                    import json
                                    parsed_template = json.loads(template_text)
                                    yaml_template = parsed_template
                                    print(f"[DEBUG] Template is in JSON format, parsed successfully")
                                except json.JSONDecodeError as json_err:
                                    print(f"[WARNING] Template is neither valid YAML nor JSON: {str(yaml_err)}, {str(json_err)}")
                                    # Build a new template from the form as fallback
                                    yaml_template = self.build_invoice2data_template()
                                    print(f"[DEBUG] Using fallback template built from form")
                        else:
                            print(f"[DEBUG] Template text is empty")
                            yaml_template = None
                    except Exception as e:
                        print(f"[ERROR] Error processing template: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        # Build a new template from the form as fallback
                        yaml_template = self.build_invoice2data_template()
                        print(f"[DEBUG] Using fallback template built from form due to error")
                else:
                    print(f"[DEBUG] No template_preview found, building template from form")
                    yaml_template = self.build_invoice2data_template()

                print(f"[DEBUG] Saving YAML template: {yaml_template is not None}")
                if yaml_template is None:
                    print(f"[DEBUG] YAML template is None or empty")

                template_id = db.save_template(
                    name=name,
                    description=description,
                    regions={},  # Empty for multi-page templates
                    column_lines={},  # Empty for multi-page templates
                    config=config,
                    template_type="multi",  # This is a multi-page template
                    page_count=len(self.pdf_document),  # Add the page count
                    page_regions=page_regions,
                    page_column_lines=page_column_lines,
                    page_configs=page_configs,  # Use the newly created page_configs list
                    json_template=yaml_template
                )

            else:
                # Single-page template
                print(f"[DEBUG] Saving single-page template")

                # Initialize config with extraction parameters
                config = self.extraction_params.copy()

                # Add additional parameters to config
                for key, value in additional_params.items():
                    config[key] = value

                # Store original coordinates if requested
                if additional_params.get('store_original_coords', True):
                    # Store original regions
                    original_regions = {}
                    for section, rects in self.regions.items():
                        original_regions[section] = []
                        for rect in rects:
                            # Check if rect is a QRect object or a dictionary with 'rect' key
                            if isinstance(rect, dict) and 'rect' in rect:
                                # This is a dictionary with a 'rect' key containing the QRect object
                                qrect = rect['rect']
                                original_regions[section].append({
                                    'x': qrect.x(),
                                    'y': qrect.y(),
                                    'width': qrect.width(),
                                    'height': qrect.height()
                                })
                            elif hasattr(rect, 'x') and callable(rect.x):
                                # This is a QRect object directly
                                original_regions[section].append({
                                    'x': rect.x(),
                                    'y': rect.y(),
                                    'width': rect.width(),
                                    'height': rect.height()
                                })
                            else:
                                # This is some other type of object, log it for debugging
                                print(f"[DEBUG] Unexpected rect type: {type(rect)}, skipping")
                    config['original_regions'] = original_regions

                    # Store original column lines
                    original_column_lines = {}
                    for section, lines in self.column_lines.items():
                        section_name = section.value if hasattr(section, 'value') else section
                        original_column_lines[section_name] = []
                        for line in lines:
                            if len(line) >= 2 and isinstance(line[0], QPoint) and isinstance(line[1], QPoint):
                                line_data = [
                                    {'x': line[0].x(), 'y': line[0].y()},
                                    {'x': line[1].x(), 'y': line[1].y()}
                                ]
                                if len(line) > 2:
                                    line_data.append(line[2])  # Add region index if present
                                original_column_lines[section_name].append(line_data)
                    config['original_column_lines'] = original_column_lines

                # Convert regions to serializable format for database
                serializable_regions = {}
                for section, rects in self.regions.items():
                    serializable_regions[section] = []
                    for rect in rects:
                        # Convert to PDF coordinates
                        scale_factors = get_scale_factors(self.pdf_path, 0)  # Single page, so page_index is 0
                        scale_x = scale_factors['scale_x']
                        scale_y = scale_factors['scale_y']
                        page_height = scale_factors['page_height']

                        # Check if rect is a QRect object or a dictionary with 'rect' key
                        if isinstance(rect, dict) and 'rect' in rect:
                            # This is a dictionary with a 'rect' key containing the QRect object
                            qrect = rect['rect']
                            # Convert QRect to PDF coordinates
                            x1 = qrect.x() * scale_x
                            y1 = page_height - (qrect.y() * scale_y)  # Convert top y1 to bottom y1
                            x2 = (qrect.x() + qrect.width()) * scale_x
                            y2 = page_height - ((qrect.y() + qrect.height()) * scale_y)  # Convert bottom y2 to bottom y2
                        elif hasattr(rect, 'x') and callable(rect.x):
                            # This is a QRect object directly
                            # Convert QRect to PDF coordinates
                            x1 = rect.x() * scale_x
                            y1 = page_height - (rect.y() * scale_y)  # Convert top y1 to bottom y1
                            x2 = (rect.x() + rect.width()) * scale_x
                            y2 = page_height - ((rect.y() + rect.height()) * scale_y)  # Convert bottom y2 to bottom y2
                        else:
                            # This is some other type of object, log it for debugging
                            print(f"[DEBUG] Unexpected rect type: {type(rect)}, skipping")
                            continue

                        # Add to serializable_regions as a dictionary with x, y, width, height
                        width = x2 - x1
                        height = y1 - y2  # Note: y1 and y2 are flipped in PDF coordinates

                        # Store as a dictionary with original coordinates
                        serializable_regions[section].append({
                            'x': x1,
                            'y': y2,  # Use bottom-left y coordinate
                            'width': width,
                            'height': height,
                            'x1': x1,
                            'y1': y1,
                            'x2': x2,
                            'y2': y2
                        })

                # Convert column lines to serializable format for database
                serializable_column_lines = {}
                for section, lines in self.column_lines.items():
                    section_name = section.value if hasattr(section, 'value') else section
                    serializable_column_lines[section_name] = []

                    # Get scale factors
                    scale_factors = get_scale_factors(self.pdf_path, 0)  # Single page, so page_index is 0
                    scale_x = scale_factors['scale_x']
                    scale_y = scale_factors['scale_y']
                    page_height = scale_factors['page_height']

                    for line in lines:
                        if len(line) >= 2 and isinstance(line[0], QPoint) and isinstance(line[1], QPoint):
                            # Convert to PDF coordinates
                            start_x = line[0].x() * scale_x
                            start_y = page_height - (line[0].y() * scale_y)
                            end_x = line[1].x() * scale_x
                            end_y = page_height - (line[1].y() * scale_y)

                            # Store as a list of dictionaries with original coordinates
                            column_data = [
                                {'x': start_x, 'y': start_y},
                                {'x': end_x, 'y': end_y}
                            ]

                            # Add region index if available
                            if len(line) > 2 and isinstance(line[2], int):
                                column_data.append(line[2])

                            serializable_column_lines[section_name].append(column_data)

                # Save template to the database
                # Get the invoice2data template JSON from the template_preview
                print(f"\n[DEBUG] Getting invoice2data template JSON for save_template_directly (single-page)")
                json_template = None

                # Get the template directly from the template_preview as YAML
                print(f"[DEBUG] Has template_preview: {hasattr(self, 'template_preview')}")
                if hasattr(self, 'template_preview') and self.template_preview:
                    try:
                        # Get the raw text from the template preview
                        template_text = self.template_preview.toPlainText()
                        print(f"[DEBUG] Template text length: {len(template_text)}")
                        print(f"[DEBUG] Template text (first 100 chars): {template_text[:100]}...")

                        # Check if the template preview contains test results
                        if template_text.strip().startswith("TEST RESULT"):
                            # This is a test result, not a valid template
                            print(f"[DEBUG] Template preview contains test results, not a valid template")
                            # Get the last valid template that was tested
                            if hasattr(self, '_last_valid_template') and self._last_valid_template:
                                yaml_template = self._last_valid_template
                                print(f"[DEBUG] Using last valid template that was tested")
                            else:
                                # Build a new template from the form
                                yaml_template = self.build_invoice2data_template()
                                print(f"[DEBUG] No last valid template found, building from form")
                        elif template_text.strip():
                            # Try to parse as YAML first (preferred format)
                            import yaml
                            try:
                                # Parse the template to validate it's proper YAML
                                parsed_template = yaml.safe_load(template_text)
                                if parsed_template and isinstance(parsed_template, dict):
                                    # Use the parsed template directly without any modifications
                                    yaml_template = parsed_template
                                    print(f"[DEBUG] Successfully parsed template as YAML: {list(yaml_template.keys())}")
                                else:
                                    print(f"[DEBUG] Template is not a valid dictionary: {type(parsed_template)}")
                                    yaml_template = self.build_invoice2data_template()
                            except yaml.YAMLError as yaml_err:
                                # If YAML parsing fails, try JSON as fallback
                                try:
                                    import json
                                    parsed_template = json.loads(template_text)
                                    yaml_template = parsed_template
                                    print(f"[DEBUG] Template is in JSON format, parsed successfully")
                                except json.JSONDecodeError as json_err:
                                    print(f"[WARNING] Template is neither valid YAML nor JSON: {str(yaml_err)}, {str(json_err)}")
                                    # Build a new template from the form as fallback
                                    yaml_template = self.build_invoice2data_template()
                                    print(f"[DEBUG] Using fallback template built from form")
                        else:
                            print(f"[DEBUG] Template text is empty")
                            yaml_template = None
                    except Exception as e:
                        print(f"[ERROR] Error processing template: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        # Build a new template from the form as fallback
                        yaml_template = self.build_invoice2data_template()
                        print(f"[DEBUG] Using fallback template built from form due to error")
                else:
                    print(f"[DEBUG] No template_preview found, building template from form")
                    yaml_template = self.build_invoice2data_template()

                print(f"[DEBUG] Saving YAML template: {yaml_template is not None}")
                if yaml_template is None:
                    print(f"[DEBUG] YAML template is None or empty")

                template_id = db.save_template(
                    name=name,
                    description=description,
                    regions=serializable_regions,  # Now using properly scaled and converted coordinates
                    column_lines=serializable_column_lines,  # Now using properly scaled and converted coordinates
                    config=config,  # Use the complete config object that includes original coordinates
                    template_type="single",
                    page_count=1,
                    json_template=yaml_template
                )

            if template_id:
                # Emit signal to notify that template was saved to database
                self.save_template_signal.emit()

                # Show a brief success message
                brief_msg = QMessageBox(self)
                brief_msg.setWindowTitle("Template Saved")
                brief_msg.setText("Template saved successfully!")
                brief_msg.setInformativeText(f"Template '{name}' has been saved. Navigating to Template Manager...")
                brief_msg.setIcon(QMessageBox.Information)
                brief_msg.setStyleSheet("QLabel { color: black; }")
                brief_msg.setStandardButtons(QMessageBox.Ok)
                brief_msg.exec()

                # Close database connection
                db.close()

                # Navigate to template manager screen
                self.navigate_to_template_manager()
                return template_id
            else:
                raise Exception("Failed to save template - no template ID returned")

        except Exception as e:
            print(f"\nError saving template: {str(e)}")
            import traceback
            traceback.print_exc()

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Error")
            msg.setText("Failed to save template")
            msg.setInformativeText(f"An error occurred: {str(e)}")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec()
            return None

    def navigate_to_template_manager(self):
        """Navigate to the template manager screen"""
        try:
            from PySide6.QtWidgets import QApplication, QStackedWidget
            import logging

            print("\n[DEBUG] Attempting to navigate to template manager")

            # Method 1: Try to find main window via direct parent hierarchy
            main_window = self.window()
            if main_window:
                print("[DEBUG] Found main window via window() method")
            else:
                # Method 2: Try searching among top-level widgets
                for widget in QApplication.topLevelWidgets():
                    # Check if it has a stacked widget child - a good hint it's our main window
                    stacked_widget = widget.findChild(QStackedWidget)
                    if stacked_widget:
                        main_window = widget
                        print("[DEBUG] Found main window by searching top-level widgets")
                        break

            if main_window:
                # Find the stacked widget
                stacked_widget = main_window.findChild(QStackedWidget)
                if stacked_widget:
                    print("[DEBUG] Found stacked widget")

                    # Find the template manager in the stacked widget
                    template_manager_index = -1
                    for i in range(stacked_widget.count()):
                        widget = stacked_widget.widget(i)
                        if widget and widget.__class__.__name__ == "TemplateManager":
                            template_manager_index = i
                            print(f"[DEBUG] Found template manager at index {i}")
                            break

                    # Navigate if we found the template manager
                    if template_manager_index >= 0:
                        # Try to refresh the template list before showing
                        template_manager = stacked_widget.widget(template_manager_index)
                        try:
                            if hasattr(template_manager, 'load_templates'):
                                template_manager.load_templates()
                                print("[DEBUG] Successfully refreshed template list")
                            else:
                                print("[WARNING] Template manager doesn't have load_templates method")
                        except Exception as refresh_error:
                            print(f"[WARNING] Error refreshing template list: {str(refresh_error)}")

                        # Navigate to the template manager screen
                        stacked_widget.setCurrentIndex(template_manager_index)
                        print(f"[DEBUG] Successfully navigated to template manager at index {template_manager_index}")
                    else:
                        print("[ERROR] Could not find template manager in the stacked widget")
                else:
                    print("[ERROR] No stacked widget available")

        except Exception as e:
            print(f"[ERROR] Failed to navigate to template manager: {str(e)}")
            import traceback
            traceback.print_exc()

    def clear_current_page(self):
        """Clear all drawings in the PDF"""
        if not self.pdf_document:
            return

        print(f"[DEBUG] Clearing all drawings in the PDF")

        # Clear all regions and column lines for all pages in multi-page mode
        if self.multi_page_mode:
            # Clear all page regions and column lines
            for page_idx in range(len(self.pdf_document)):
                self.page_regions[page_idx] = {'header': [], 'items': [], 'summary': []}
                self.page_column_lines[page_idx] = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}

            # Also clear the current page's regions and column lines in memory
            self.regions = {'header': [], 'items': [], 'summary': []}
            self.column_lines = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}

            print(f"[DEBUG] Cleared drawings for all {len(self.pdf_document)} pages")
        else:
            # Clear regions and column lines for single-page mode
            self.regions = {'header': [], 'items': [], 'summary': []}
            self.column_lines = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}
            print(f"[DEBUG] Cleared all drawings")

        # Reset drawing state
        self.current_region_type = None
        self.drawing_column = False

        # Reset cursor to default
        if self.pdf_label:
            self.pdf_label.setCursor(Qt.ArrowCursor)

        # Uncheck all buttons
        self.header_btn.setChecked(False)
        self.items_btn.setChecked(False)
        self.summary_btn.setChecked(False)
        self.column_btn.setChecked(False)

        # Clear cached extraction data for all pages
        self._cached_extraction_data = {
            'header': [],
            'items': [],
            'summary': []
        }
        self._last_extraction_state = None

        # Clear all pages data
        if hasattr(self, '_all_pages_data') and self._all_pages_data is not None:
            # Only clear the current page data in multi-page mode
            if self.multi_page_mode:
                if self.current_page_index < len(self._all_pages_data):
                    self._all_pages_data[self.current_page_index] = None
                    print(f"[DEBUG] Cleared extraction data for page {self.current_page_index + 1}")
            else:
                # In single-page mode, clear all pages
                self._all_pages_data = [None] * len(self.pdf_document)
                print(f"[DEBUG] Cleared extraction data for all pages")

        # Update the JSON tree
        self.update_json_tree(self._cached_extraction_data)

        # Reset the JSON designer screen
        self.reset_json_designer()

        # Update the display
        self.pdf_label.update()

    def reset_screen(self):
        """Reset the screen to its initial state, clear cache, and clean temp directory"""
        print(f"[DEBUG] Resetting screen")

        # Clear drawings for the current page
        self.clear_current_page()

        # Reset the JSON designer screen
        self.reset_json_designer()

        # Clear memory cache
        import gc
        gc.collect()

        # Clean temp directory
        try:
            import os
            import shutil

            # Get the temp directory using path_helper
            if PATH_HELPER_AVAILABLE:
                temp_dir = path_helper.ensure_directory("temp")
            else:
                temp_dir = os.path.abspath("temp")
                os.makedirs(temp_dir, exist_ok=True)

            # Check if the directory exists
            if os.path.exists(temp_dir):
                # Remove all files in the directory
                for file in os.listdir(temp_dir):
                    file_path = os.path.join(temp_dir, file)
                    try:
                        if os.path.isfile(file_path):
                            os.unlink(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print(f"[ERROR] Failed to remove {file_path}: {str(e)}")

                print(f"[DEBUG] Cleaned temp directory: {temp_dir}")
            else:
                print(f"[DEBUG] Temp directory does not exist: {temp_dir}")

        except Exception as e:
            print(f"[ERROR] Failed to clean temp directory: {str(e)}")
            import traceback
            traceback.print_exc()

        # Clear extraction cache
        if hasattr(self, 'pdf_path') and self.pdf_path:
            clear_extraction_cache_for_pdf(self.pdf_path)
            print(f"[DEBUG] Cleared extraction cache for PDF: {self.pdf_path}")

        # Show success message
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.information(
            self,
            "Reset Complete",
            "Screen has been reset, cache cleared, and temp files removed."
        )

    def reset_json_designer(self):
        """Reset the JSON designer screen to its default state"""
        print(f"[DEBUG] Resetting JSON designer screen")

        # Reset the invoice2data template to default values using factory
        base_template = TemplateFactory.create_default_invoice_template("", "INR")
        self.invoice2data_template = {
            "issuer": base_template["issuer"],
            "fields": {},
            "lines": {
                "start": "",
                "end": "",
                "first_line": [],
                "line": "",
                "types": {}
            },
            "keywords": base_template["keywords"],
            "options": {
                **base_template["options"],
                "date_formats": [],  # Empty by default, user can add if needed
                "remove_whitespace": False,
                "remove_accents": False,
                "lowercase": False,
                "replace": []
            }
        }

        # If the invoice2data editor exists, reset its fields
        if hasattr(self, 'invoice2data_editor') and self.invoice2data_editor:
            # Reset the template preview if it exists
            if hasattr(self, 'template_preview') and self.template_preview:
                # Format the template as YAML
                try:
                    import yaml
                    formatted_yaml = yaml.safe_dump(
                        self.invoice2data_template,
                        default_flow_style=False,
                        allow_unicode=True,
                        sort_keys=False,
                        indent=2,
                        width=80
                    )
                    self.template_preview.setText(formatted_yaml)
                    print(f"[DEBUG] Reset template preview with default template")
                except Exception as e:
                    print(f"[ERROR] Failed to reset template preview: {str(e)}")

            # Reset form fields if they exist
            try:
                # General tab fields
                if hasattr(self, 'issuer_input'):
                    self.issuer_input.setText("")

                if hasattr(self, 'priority_input'):
                    self.priority_input.setValue(100)

                if hasattr(self, 'keywords_list'):
                    self.keywords_list.clear()

                if hasattr(self, 'exclude_keywords_list'):
                    self.exclude_keywords_list.clear()

                # Options
                if hasattr(self, 'currency_input'):
                    self.currency_input.setText("INR")

                if hasattr(self, 'decimal_separator_input'):
                    self.decimal_separator_input.setText(".")

                if hasattr(self, 'language_input'):
                    self.language_input.setCurrentText("en")

                if hasattr(self, 'date_formats_list'):
                    self.date_formats_list.clear()
                    for format_str in self.invoice2data_template["options"]["date_formats"]:
                        self.date_formats_list.addItem(format_str)

                if hasattr(self, 'remove_whitespace_checkbox'):
                    self.remove_whitespace_checkbox.setChecked(False)

                if hasattr(self, 'remove_accents_checkbox'):
                    self.remove_accents_checkbox.setChecked(False)

                if hasattr(self, 'lowercase_checkbox'):
                    self.lowercase_checkbox.setChecked(False)

                # Replace patterns
                if hasattr(self, 'replace_table'):
                    while self.replace_table.rowCount() > 0:
                        self.replace_table.removeRow(0)

                # Lines
                if hasattr(self, 'line_start_input'):
                    self.line_start_input.clear()

                if hasattr(self, 'line_end_input'):
                    self.line_end_input.clear()

                if hasattr(self, 'line_pattern_input'):
                    self.line_pattern_input.clear()

                if hasattr(self, 'skip_line_input'):
                    self.skip_line_input.clear()

                if hasattr(self, 'firstline_list'):
                    self.firstline_list.clear()

                # Line types
                if hasattr(self, 'line_types_table'):
                    while self.line_types_table.rowCount() > 0:
                        self.line_types_table.removeRow(0)

                # Fields
                if hasattr(self, 'fields_table'):
                    while self.fields_table.rowCount() > 0:
                        self.fields_table.removeRow(0)

                print(f"[DEBUG] Successfully reset all JSON designer form fields")
            except Exception as e:
                print(f"[ERROR] Error resetting JSON designer form fields: {str(e)}")
                import traceback
                traceback.print_exc()

    def create_invoice2data_editor(self):
        """Create the invoice2data template editor widget"""
        editor_widget = QWidget()
        editor_layout = QVBoxLayout(editor_widget)

        # # Title
        # title = QLabel("Invoice2Data Template Creator")
        # title.setFont(QFont("Arial", 16, QFont.Bold))
        # title.setAlignment(Qt.AlignCenter)
        # editor_layout.addWidget(title)

        # # Description


        # Create tab widget for different sections
        tab_widget = QTabWidget()

        # Populate the form with default values from the template
        self.populate_form_from_template()

        # General tab
        general_tab = QWidget()
        general_layout = QFormLayout(general_tab)

        # Issuer field
        self.issuer_input = QLineEdit()
        general_layout.addRow("Issuer:", self.issuer_input)

        # Priority field
        self.priority_input = QSpinBox()
        self.priority_input.setRange(0, 1000)
        self.priority_input.setValue(100)  # Default value of 100
        self.priority_input.setToolTip("Priority of this template (higher values = higher priority)")
        general_layout.addRow("Priority:", self.priority_input)

        # Keywords field (list widget with add/remove buttons)
        keywords_container = QWidget()
        keywords_layout = QVBoxLayout(keywords_container)
        keywords_layout.setContentsMargins(0, 0, 0, 0)

        self.keywords_list = QListWidget()
        keywords_layout.addWidget(self.keywords_list)

        keywords_buttons = QHBoxLayout()
        add_keyword_btn = QPushButton("Add Keyword")
        add_keyword_btn.clicked.connect(self.add_keyword)
        remove_keyword_btn = QPushButton("Remove Selected")
        remove_keyword_btn.clicked.connect(self.remove_keyword)
        keywords_buttons.addWidget(add_keyword_btn)
        keywords_buttons.addWidget(remove_keyword_btn)
        keywords_layout.addLayout(keywords_buttons)

        general_layout.addRow("Keywords:", keywords_container)

        # Exclude Keywords field (list widget with add/remove buttons)
        exclude_keywords_container = QWidget()
        exclude_keywords_layout = QVBoxLayout(exclude_keywords_container)
        exclude_keywords_layout.setContentsMargins(0, 0, 0, 0)

        self.exclude_keywords_list = QListWidget()
        exclude_keywords_layout.addWidget(self.exclude_keywords_list)

        exclude_keywords_buttons = QHBoxLayout()
        add_exclude_keyword_btn = QPushButton("Add Exclude Keyword")
        add_exclude_keyword_btn.clicked.connect(self.add_exclude_keyword)
        remove_exclude_keyword_btn = QPushButton("Remove Selected")
        remove_exclude_keyword_btn.clicked.connect(self.remove_exclude_keyword)
        exclude_keywords_buttons.addWidget(add_exclude_keyword_btn)
        exclude_keywords_buttons.addWidget(remove_exclude_keyword_btn)
        exclude_keywords_layout.addLayout(exclude_keywords_buttons)

        general_layout.addRow("Exclude Keywords:", exclude_keywords_container)

        # Options section
        options_group = QGroupBox("Options")
        options_layout = QFormLayout(options_group)

        self.currency_input = QLineEdit("INR")
        options_layout.addRow("Currency:", self.currency_input)

        self.decimal_separator_input = QLineEdit(".")
        options_layout.addRow("Decimal Separator:", self.decimal_separator_input)

        # Language selection
        self.language_input = QComboBox()
        self.language_input.addItems(["en", "nl", "de", "fr", "es", "it"])
        self.language_input.setCurrentText("en")  # Default to English
        options_layout.addRow("Language:", self.language_input)

        # Date formats (list widget with add/remove buttons)
        date_formats_container = QWidget()
        date_formats_layout = QVBoxLayout(date_formats_container)
        date_formats_layout.setContentsMargins(0, 0, 0, 0)

        self.date_formats_list = QListWidget()
        date_formats_layout.addWidget(self.date_formats_list)

        date_formats_buttons = QHBoxLayout()
        add_date_format_btn = QPushButton("Add Format")
        add_date_format_btn.clicked.connect(self.add_date_format)
        remove_date_format_btn = QPushButton("Remove Selected")
        remove_date_format_btn.clicked.connect(self.remove_date_format)
        date_formats_buttons.addWidget(add_date_format_btn)
        date_formats_buttons.addWidget(remove_date_format_btn)
        date_formats_layout.addLayout(date_formats_buttons)

        options_layout.addRow("Date Formats:", date_formats_container)

        # Additional options (checkboxes)
        additional_options_container = QWidget()
        additional_options_layout = QVBoxLayout(additional_options_container)
        additional_options_layout.setContentsMargins(0, 0, 0, 0)

        self.remove_whitespace_checkbox = QCheckBox("Remove Whitespace")
        self.remove_whitespace_checkbox.setToolTip("Ignore any spaces. Often makes regex easier to write.")
        additional_options_layout.addWidget(self.remove_whitespace_checkbox)

        self.remove_accents_checkbox = QCheckBox("Remove Accents")
        self.remove_accents_checkbox.setToolTip("Useful for international invoices. Saves you from putting accents in your regular expressions.")
        additional_options_layout.addWidget(self.remove_accents_checkbox)

        self.lowercase_checkbox = QCheckBox("Lowercase")
        self.lowercase_checkbox.setToolTip("Convert all text to lowercase before matching.")
        additional_options_layout.addWidget(self.lowercase_checkbox)

        options_layout.addRow("Text Processing:", additional_options_container)

        # Replace patterns table
        replace_container = QWidget()
        replace_layout = QVBoxLayout(replace_container)
        replace_layout.setContentsMargins(0, 0, 0, 0)

        # Table for replace patterns
        self.replace_table = QTableWidget(0, 2)
        self.replace_table.setHorizontalHeaderLabels(["Find", "Replace"])
        self.replace_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.replace_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        replace_layout.addWidget(self.replace_table)

        # Buttons for replace patterns
        replace_buttons = QHBoxLayout()
        add_replace_btn = QPushButton("Add Pattern")
        add_replace_btn.clicked.connect(self.add_replace_pattern)
        remove_replace_btn = QPushButton("Remove Selected")
        remove_replace_btn.clicked.connect(self.remove_replace_pattern)
        replace_buttons.addWidget(add_replace_btn)
        replace_buttons.addWidget(remove_replace_btn)
        replace_layout.addLayout(replace_buttons)

        options_layout.addRow("Replace Patterns:", replace_container)

        general_layout.addRow("", options_group)

        # Preview tab
        preview_tab = QWidget()
        preview_layout = QVBoxLayout(preview_tab)

        self.template_preview = QTextEdit()
        self.template_preview.setReadOnly(False)  # Make it editable
        self.template_preview.setFont(QFont("Courier New", 10))  # Use monospace font for better JSON editing
        self.template_preview.setLineWrapMode(QTextEdit.NoWrap)  # Disable line wrapping for better JSON editing
        preview_layout.addWidget(self.template_preview)

        # Add a label to indicate the preview is editable
        edit_label = QLabel("The template preview is editable. Make changes and click 'Test Template' to test with your edits.")
        edit_label.setStyleSheet("color: #0066cc; font-style: italic;")
        preview_layout.addWidget(edit_label)

        preview_buttons = QHBoxLayout()

        update_preview_btn = QPushButton("Update Preview")
        update_preview_btn.setToolTip("Update the preview with current form values")
        update_preview_btn.clicked.connect(self.update_template_preview)

        apply_changes_btn = QPushButton("Apply Changes")
        apply_changes_btn.setToolTip("Apply changes from the edited preview to the form")
        apply_changes_btn.clicked.connect(self.apply_template_changes)
        apply_changes_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        test_template_btn = QPushButton("Test Template")
        test_template_btn.setToolTip("Test the current template against the loaded PDF")
        test_template_btn.clicked.connect(self.test_invoice2data_template)

        preview_buttons.addWidget(update_preview_btn)
        preview_buttons.addWidget(apply_changes_btn)
        preview_buttons.addWidget(test_template_btn)
        preview_layout.addLayout(preview_buttons)

        # Add tabs to tab widget
        tab_widget.addTab(general_tab, "General")
        tab_widget.addTab(preview_tab, "Preview & Test")

        editor_layout.addWidget(tab_widget)

        # No bottom buttons - using the ones in the Preview tab instead

        return editor_widget


    def on_top_tab_changed(self, index):
        """Handle top tab changes to synchronize with bottom tabs"""
        # Skip if it's the Metadata tab (index 0)
        if index == 0:
            # Hide the bottom tabs when Metadata is selected
            self.bottom_tabs.setVisible(False)
            return
        else:
            # Show the bottom tabs for other selections
            self.bottom_tabs.setVisible(True)

        # Make sure we have a template loaded and fields are populated
        if hasattr(self, 'toggle_json_designer_btn') and self.toggle_json_designer_btn.isChecked():
            if not hasattr(self, 'invoice2data_template') or self.invoice2data_template is None:
                self.initialize_invoice2data_template()
                print(f"[DEBUG] Initialized invoice2data template when tab changed in top tab change")
            elif hasattr(self, 'header_fields_table') and self.header_fields_table.rowCount() == 0:
                # Populate the form with the template values
                self.populate_form_from_template()
                print(f"[DEBUG] Populated form from template when tab changed in top tab change")
        else:
            # Hide all editors if JSON designer is toggled off
            if hasattr(self, 'fields_editor'):
                self.fields_editor.hide()
            if hasattr(self, 'tables_editor'):
                self.tables_editor.hide()
            if hasattr(self, 'lines_editor'):
                self.lines_editor.hide()
            if hasattr(self, 'tax_lines_editor'):
                self.tax_lines_editor.hide()

    def on_bottom_tab_changed(self, index):
        """Handle changes to the bottom tab widget"""
        # Make sure we have a template loaded and fields are populated
        if hasattr(self, 'toggle_json_designer_btn') and self.toggle_json_designer_btn.isChecked():
            if not hasattr(self, 'invoice2data_template') or self.invoice2data_template is None:
                self.initialize_invoice2data_template()
                print(f"[DEBUG] Initialized invoice2data template when bottom tab changed")
            elif hasattr(self, 'header_fields_table') and self.header_fields_table.rowCount() == 0:
                # Populate the form with the template values
                self.populate_form_from_template()
                print(f"[DEBUG] Populated form from template when bottom tab changed")

        # Log the tab change
        print(f"[DEBUG] Bottom tab changed to index {index}")

    def show_tree_context_menu(self, position):
        """Show context menu for the extraction results view"""
        menu = QMenu()
        copy_action = menu.addAction("Copy")
        copy_all_action = menu.addAction("Copy All")

        # Get the sender widget
        sender = self.sender()

        # Get the action that was clicked
        action = menu.exec_(sender.mapToGlobal(position))

        if action == copy_action:
            # Copy selected text
            if isinstance(sender, QTextEdit):
                cursor = sender.textCursor()
                selected_text = cursor.selectedText()
                if selected_text:
                    QApplication.clipboard().setText(selected_text)
                    return

        if action == copy_all_action:
            # Copy all text
            if isinstance(sender, QTextEdit):
                QApplication.clipboard().setText(sender.toPlainText())
            else:
                # Fallback to old behavior
                self.copy_all_data_to_clipboard()

    # on_tree_item_clicked method has been removed as it was unused

    def on_regex_pattern_changed(self):
        """Handle changes to regex patterns in the form"""
        # This method is called when any regex pattern input is changed
        # It can be used to provide real-time validation or feedback
        sender = self.sender()
        if sender:
            print(f"[DEBUG] Regex pattern changed in {sender.objectName()}")

    def populate_form_from_template(self):
        """Populate the form with default values from the template"""
        try:
            # Populate form with default values from the template
            if self.invoice2data_template:
                # Set priority if available
                if 'priority' in self.invoice2data_template:
                    try:
                        priority_value = int(self.invoice2data_template['priority'])
                        self.priority_input.setValue(priority_value)
                    except (ValueError, TypeError):
                        print(f"[WARNING] Invalid priority value in template: {self.invoice2data_template.get('priority')}, using default 100")
                        self.priority_input.setValue(100)
                else:
                    # Use default priority
                    self.priority_input.setValue(100)

                # Set options if available
                if 'options' in self.invoice2data_template:
                    options = self.invoice2data_template['options']

                # Populate header fields table with template fields
                if hasattr(self, 'header_fields_table') and 'fields' in self.invoice2data_template:
                    fields = self.invoice2data_template['fields']

                    # Clear existing fields
                    while self.header_fields_table.rowCount() > 0:
                        self.header_fields_table.removeRow(0)

                    # Add fields from template
                    for field_name, field_data in fields.items():
                        row = self.header_fields_table.rowCount()
                        self.header_fields_table.insertRow(row)

                        # Set field name
                        self.header_fields_table.setItem(row, 0, QTableWidgetItem(field_name))

                        # Set regex pattern
                        if isinstance(field_data, dict) and 'regex' in field_data:
                            regex_pattern = field_data['regex']
                        else:
                            # Simple field format (just a regex pattern)
                            regex_pattern = str(field_data)
                        self.header_fields_table.setItem(row, 1, QTableWidgetItem(regex_pattern))

                        # Set field type
                        type_combo = QComboBox()
                        type_combo.addItems(["string", "date", "float", "int"])
                        if isinstance(field_data, dict) and 'type' in field_data:
                            field_type = field_data['type']
                            if field_type in ["string", "date", "float", "int"]:
                                type_combo.setCurrentText(field_type)
                        self.header_fields_table.setCellWidget(row, 2, type_combo)

                    print(f"[DEBUG] Populated header fields table with {self.header_fields_table.rowCount()} fields from template")

                # We no longer need to populate a separate summary fields table since we're using a common tab structure

                # Populate replace patterns table
                if hasattr(self, 'replace_table') and 'replace' in options and options['replace']:
                    # Clear existing patterns
                    while self.replace_table.rowCount() > 0:
                        self.replace_table.removeRow(0)

                    # Add patterns from template
                    for pattern in options['replace']:
                        if isinstance(pattern, list) and len(pattern) >= 2:
                            row = self.replace_table.rowCount()
                            self.replace_table.insertRow(row)
                            self.replace_table.setItem(row, 0, QTableWidgetItem(pattern[0]))
                            self.replace_table.setItem(row, 1, QTableWidgetItem(pattern[1]))

                # Populate date formats list
                if hasattr(self, 'date_formats_list') and 'date_formats' in options and options['date_formats']:
                    # Clear existing date formats
                    self.date_formats_list.clear()

                    # Add date formats from template
                    for format_str in options['date_formats']:
                        self.date_formats_list.addItem(format_str)

                # Set additional options
                if hasattr(self, 'remove_whitespace_checkbox'):
                    self.remove_whitespace_checkbox.setChecked(options.get('remove_whitespace', False))
                if hasattr(self, 'remove_accents_checkbox'):
                    self.remove_accents_checkbox.setChecked(options.get('remove_accents', False))
                if hasattr(self, 'lowercase_checkbox'):
                    self.lowercase_checkbox.setChecked(options.get('lowercase', False))
        except Exception as e:
            print(f"[ERROR] Failed to populate form from template: {str(e)}")
            import traceback
            traceback.print_exc()

    # Methods for Fields tab
    def add_field(self):
        """Add a new field to the fields table"""
        row = self.header_fields_table.rowCount()
        self.header_fields_table.insertRow(row)

        # Set default field name
        self.header_fields_table.setItem(row, 0, QTableWidgetItem("field_name"))

        # Set default regex pattern
        self.header_fields_table.setItem(row, 1, QTableWidgetItem("(.+)"))

        # Create a combo box for field type
        type_combo = QComboBox()
        type_combo.addItems(["string", "date", "float", "int"])
        self.header_fields_table.setCellWidget(row, 2, type_combo)

        print(f"[DEBUG] Added new field at row {row}")

    def remove_field(self):
        """Remove the selected field from the fields table"""
        selected_rows = self.header_fields_table.selectedIndexes()
        if not selected_rows:
            return

        # Get unique row indices and sort them in descending order
        rows = sorted(set(index.row() for index in selected_rows), reverse=True)

        # Remove rows from bottom to top to avoid index shifting
        for row in rows:
            self.header_fields_table.removeRow(row)
            print(f"[DEBUG] Removed field at row {row}")

    # We no longer need separate methods for summary fields since we're using a common tab structure
    # The add_field and remove_field methods will be used for all fields

    # Methods for Lines tab
    def add_firstline_pattern(self):
        """Add a new first line pattern to the list"""
        text, ok = QInputDialog.getText(self, "Add First Line Pattern", "Enter regex pattern:")
        if ok and text:
            self.firstline_list.addItem(text)
            print(f"[DEBUG] Added first line pattern: {text}")

    def remove_firstline_pattern(self):
        """Remove the selected first line pattern from the list"""
        selected_items = self.firstline_list.selectedItems()
        for item in selected_items:
            row = self.firstline_list.row(item)
            self.firstline_list.takeItem(row)
            print(f"[DEBUG] Removed first line pattern at row {row}")

    def add_line_type(self):
        """Add a new line type to the line types table"""
        row = self.line_types_table.rowCount()
        self.line_types_table.insertRow(row)

        # Set default field name
        self.line_types_table.setItem(row, 0, QTableWidgetItem("field_name"))

        # Create a combo box for field type
        type_combo = QComboBox()
        type_combo.addItems(["string", "date", "float", "int"])
        self.line_types_table.setCellWidget(row, 1, type_combo)

        print(f"[DEBUG] Added new line type at row {row}")

    def remove_line_type(self):
        """Remove the selected line type from the line types table"""
        selected_rows = self.line_types_table.selectedIndexes()
        if not selected_rows:
            return

        # Get unique row indices and sort them in descending order
        rows = sorted(set(index.row() for index in selected_rows), reverse=True)

        # Remove rows from bottom to top to avoid index shifting
        for row in rows:
            self.line_types_table.removeRow(row)
            print(f"[DEBUG] Removed line type at row {row}")

    # Methods for Tables tab
    def add_table_definition(self):
        """Add a new table definition to the tables definition table"""
        row = self.tables_definition_table.rowCount()
        self.tables_definition_table.insertRow(row)

        # Set default values
        self.tables_definition_table.setItem(row, 0, QTableWidgetItem("Start Pattern"))
        self.tables_definition_table.setItem(row, 1, QTableWidgetItem("End Pattern"))
        self.tables_definition_table.setItem(row, 2, QTableWidgetItem("Body Pattern"))

        print(f"[DEBUG] Added new table definition at row {row}")

    def remove_table_definition(self):
        """Remove the selected table definition from the tables definition table"""
        selected_rows = self.tables_definition_table.selectedIndexes()
        if not selected_rows:
            return

        # Get unique row indices and sort them in descending order
        rows = sorted(set(index.row() for index in selected_rows), reverse=True)

        # Remove rows from bottom to top to avoid index shifting
        for row in rows:
            self.tables_definition_table.removeRow(row)
            print(f"[DEBUG] Removed table definition at row {row}")

    # Methods for Tax Lines tab
    def add_tax_line_type(self):
        """Add a new tax line type to the tax line types table"""
        row = self.tax_line_types_table.rowCount()
        self.tax_line_types_table.insertRow(row)

        # Set default field name
        self.tax_line_types_table.setItem(row, 0, QTableWidgetItem("tax_field_name"))

        # Create a combo box for field type
        type_combo = QComboBox()
        type_combo.addItems(["string", "float", "int"])
        self.tax_line_types_table.setCellWidget(row, 1, type_combo)

        print(f"[DEBUG] Added new tax line type at row {row}")

    def remove_tax_line_type(self):
        """Remove the selected tax line type from the tax line types table"""
        selected_rows = self.tax_line_types_table.selectedIndexes()
        if not selected_rows:
            return

        # Get unique row indices and sort them in descending order
        rows = sorted(set(index.row() for index in selected_rows), reverse=True)

        # Remove rows from bottom to top to avoid index shifting
        for row in rows:
            self.tax_line_types_table.removeRow(row)
            print(f"[DEBUG] Removed tax line type at row {row}")

    # Methods for Keywords in General tab (see below for implementations)

    # Methods for Replace Patterns in General tab
    def add_replace_pattern(self):
        """Add a new replace pattern to the replace patterns table"""
        row = self.replace_table.rowCount()
        self.replace_table.insertRow(row)

        # Set default values
        self.replace_table.setItem(row, 0, QTableWidgetItem("Find"))
        self.replace_table.setItem(row, 1, QTableWidgetItem("Replace"))

        print(f"[DEBUG] Added new replace pattern at row {row}")

    def remove_replace_pattern(self):
        """Remove the selected replace pattern from the replace patterns table"""
        selected_rows = self.replace_table.selectedIndexes()
        if not selected_rows:
            return

        # Get unique row indices and sort them in descending order
        rows = sorted(set(index.row() for index in selected_rows), reverse=True)

        # Remove rows from bottom to top to avoid index shifting
        for row in rows:
            self.replace_table.removeRow(row)
            print(f"[DEBUG] Removed replace pattern at row {row}")

    # Methods for Template Preview
    def update_template_preview(self):
        """Update the template preview with the current form values"""
        try:
            # Build the template from the form values
            template = self.build_invoice2data_template()

            # Convert to YAML format
            import yaml
            yaml_text = yaml.dump(template, default_flow_style=False, sort_keys=False)

            # Update the preview
            self.template_preview.setText(yaml_text)
            print(f"[DEBUG] Updated template preview")
        except Exception as e:
            print(f"[ERROR] Failed to update template preview: {str(e)}")

    def apply_template_changes(self):
        """Apply changes from the edited preview to the form"""
        try:
            # Get the text from the preview
            text = self.template_preview.toPlainText()

            # Parse the YAML
            import yaml
            template = yaml.safe_load(text)

            # Store the template
            self.invoice2data_template = template

            # Populate the form with the template values
            self.populate_form_from_template()

            print(f"[DEBUG] Applied template changes to form")
        except Exception as e:
            print(f"[ERROR] Failed to apply template changes: {str(e)}")

    def test_invoice2data_template(self):
        """Test the current template against the loaded PDF"""
        try:
            # Get the text from the preview
            text = self.template_preview.toPlainText()

            # Parse the YAML
            import yaml
            template = yaml.safe_load(text)

            # Store the template for future use
            self._last_valid_template = template

            # Test the template against the loaded PDF
            if not self.pdf_path:
                QMessageBox.warning(self, "No PDF Loaded", "Please load a PDF before testing the template.")
                return

            # Get the current extraction data
            current_data = self._get_current_json_data()
            if not current_data:
                QMessageBox.warning(self, "No Extraction Data", "Please extract data from the PDF before testing the template.")
                return

            # Process with invoice2data
            from invoice_processing_utils import process_with_invoice2data
            result = process_with_invoice2data(self.pdf_path, template, current_data)

            # Show the result
            if result:
                # Format the result as YAML
                result_text = yaml.dump(result, default_flow_style=False, sort_keys=False)

                # Update the preview with the test result
                self.template_preview.setText(f"TEST RESULT:\n\n{result_text}")
                print(f"[DEBUG] Template test successful")
            else:
                # Show error message
                self.template_preview.setText("TEST RESULT:\n\nNo data extracted. The template did not match the PDF.")
                print(f"[DEBUG] Template test failed - no data extracted")
        except Exception as e:
            print(f"[ERROR] Failed to test template: {str(e)}")
            import traceback
            traceback.print_exc()

            # Show error message
            self.template_preview.setText(f"TEST RESULT:\n\nError testing template: {str(e)}")

    def build_invoice2data_template(self):
        """Build the invoice2data template from the form values"""
        try:
            # Get issuer
            issuer = self.issuer_input.text().strip() or "Unknown Issuer"

            # Get fields
            fields_data = {}
            for row in range(self.fields_table.rowCount()):
                field_name = self.fields_table.item(row, 0).text().strip()
                regex_pattern = self.fields_table.item(row, 1).text().strip()
                field_type = self.fields_table.cellWidget(row, 2).currentText()

                # Create field data
                field_data = {
                    'parser': 'regex',
                    'regex': regex_pattern,
                    'type': field_type
                }

                # Add date formats if field type is date
                if field_type == 'date':
                    date_formats = []
                    for i in range(self.date_formats_list.count()):
                        date_formats.append(self.date_formats_list.item(i).text())
                    if date_formats:
                        field_data['formats'] = date_formats

                # Add field to fields data
                fields_data[field_name] = field_data

            # Get keywords
            keywords = []
            for i in range(self.keywords_list.count()):
                keywords.append(self.keywords_list.item(i).text())

            # Get exclude keywords
            exclude_keywords = []
            for i in range(self.exclude_keywords_list.count()):
                exclude_keywords.append(self.exclude_keywords_list.item(i).text())

            # Get options
            options_data = {
                'currency': self.currency_input.text().strip() or "USD",
                'decimal_separator': self.decimal_separator_input.text().strip() or ".",
                'languages': [self.language_input.currentText()],
                'remove_whitespace': self.remove_whitespace_checkbox.isChecked(),
                'remove_accents': self.remove_accents_checkbox.isChecked(),
                'lowercase': self.lowercase_checkbox.isChecked()
            }

            # Get date formats
            date_formats = []
            for i in range(self.date_formats_list.count()):
                date_formats.append(self.date_formats_list.item(i).text())
            if date_formats:
                options_data['date_formats'] = date_formats

            # Get replace patterns
            replace_patterns = []
            for row in range(self.replace_table.rowCount()):
                find_text = self.replace_table.item(row, 0).text()
                replace_text = self.replace_table.item(row, 1).text()
                replace_patterns.append([find_text, replace_text])
            if replace_patterns:
                options_data['replace'] = replace_patterns

            # Build the template
            from invoice_processing_utils import build_invoice2data_template
            template = build_invoice2data_template(
                issuer=issuer,
                fields_data=fields_data,
                options_data=options_data,
                keywords=keywords,
                exclude_keywords=exclude_keywords
            )

            return template
        except Exception as e:
            print(f"[ERROR] Failed to build template: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def populate_template_from_extracted_data(self):
        """Populate the invoice2data template fields from extracted data"""
        try:
            # Get the current JSON data
            current_data = self._get_current_json_data()
            if not current_data:
                return

            # Try to extract issuer from metadata
            if 'metadata' in current_data and 'template_name' in current_data['metadata']:
                self.issuer_input.setText(current_data['metadata']['template_name'])

            # Clear existing keywords
            self.keywords_list.clear()

            # Add the issuer name as a keyword if available
            if 'metadata' in current_data and 'template_name' in current_data['metadata']:
                issuer = current_data['metadata']['template_name']
                if issuer:
                    self.keywords_list.addItem(issuer)

            # Add the PDF filename as keyword if available
            if 'metadata' in current_data and 'filename' in current_data['metadata']:
                filename = current_data['metadata']['filename']
                if filename:
                    self.keywords_list.addItem(filename)

            # Extract fields from header section
            if 'header' in current_data:
                header_data = current_data['header']
                self.extract_fields_from_section(header_data)
            elif 'InvoiceHeader' in current_data:
                header_data = current_data['InvoiceHeader']
                self.extract_fields_from_section(header_data)

            # Extract line items pattern if available
            if 'items' in current_data and current_data['items']:
                items_data = current_data['items']
                self.extract_line_items_pattern(items_data)
            elif 'InvoiceItems' in current_data and current_data['InvoiceItems']:
                items_data = current_data['InvoiceItems']
                self.extract_line_items_pattern(items_data)

        except Exception as e:
            print(f"[ERROR] Failed to populate template from extracted data: {str(e)}")

    def extract_fields_from_section(self, section_data):
        """Extract fields from a section of the extracted data"""
        # Process different data structures
        if isinstance(section_data, dict):
            for key, value in section_data.items():
                if isinstance(value, dict):
                    # Recursively process nested dictionaries
                    self.extract_fields_from_section(value)
                elif isinstance(value, str) and key not in ['page_1', 'table_0', 'table_1', 'table_2']:
                    # Add as a field
                    self.add_field_to_template(key, value)
        elif isinstance(section_data, list):
            for item in section_data:
                if isinstance(item, dict):
                    self.extract_fields_from_section(item)

    def add_field_to_template(self, field_name, sample_value):
        """Add a field to the template based on a sample value"""
        # Skip if field already exists
        for row in range(self.fields_table.rowCount()):
            if self.fields_table.item(row, 0).text() == field_name:
                return

        # Add new row
        row_count = self.fields_table.rowCount()
        self.fields_table.insertRow(row_count)

        # Set field name
        self.fields_table.setItem(row_count, 0, QTableWidgetItem(field_name))

        # Create a regex pattern based on the sample value
        pattern = self.create_regex_pattern(field_name, sample_value)
        self.fields_table.setItem(row_count, 1, QTableWidgetItem(pattern))

        # Determine field type
        type_combo = QComboBox()
        type_combo.addItems(["string", "date", "float", "int"])

        # Try to guess the type
        if re.match(r'^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$', sample_value) or \
           re.match(r'^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}$', sample_value):
            type_combo.setCurrentText("date")
        elif re.match(r'^\d+\.\d+$', sample_value):
            type_combo.setCurrentText("float")
        elif re.match(r'^\d+$', sample_value):
            type_combo.setCurrentText("int")

        self.fields_table.setCellWidget(row_count, 2, type_combo)

    def create_regex_pattern(self, field_name, sample_value):
        """Create a regex pattern based on the field name and sample value"""
        # Escape special regex characters in the sample value
        escaped_value = re.escape(sample_value)

        # Create patterns based on field name
        if 'date' in field_name.lower():
            # For dates, create a more generic pattern
            if re.match(r'^\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}$', sample_value):
                return f"{field_name.replace('_', ' ')}\\s*[:\\s]\\s*(\\d{{1,2}}[/.-]\\d{{1,2}}[/.-]\\d{{2,4}})"
            elif re.match(r'^\d{4}[/.-]\d{1,2}[/.-]\d{1,2}$', sample_value):
                return f"{field_name.replace('_', ' ')}\\s*[:\\s]\\s*(\\d{{4}}[/.-]\\d{{1,2}}[/.-]\\d{{1,2}})"
        elif 'amount' in field_name.lower() or 'total' in field_name.lower() or 'price' in field_name.lower():
            # For monetary values
            return f"{field_name.replace('_', ' ')}\\s*[:\\s]\\s*([\\d,]+\\.\\d+)"
        elif 'number' in field_name.lower() or 'invoice' in field_name.lower() or 'id' in field_name.lower():
            # For invoice numbers or IDs
            return f"{field_name.replace('_', ' ')}\\s*[:\\s]\\s*([\\w\\d\\-]+)"

        # Default pattern - use escaped_value to avoid regex special characters
        return f"{field_name.replace('_', ' ')}\\s*[:\\s]\\s*({escaped_value})"

    def extract_line_items_pattern(self, items_data):
        """Extract line items pattern from the items section"""
        # Try to identify the start and end patterns for line items
        if isinstance(items_data, list) and len(items_data) > 0:
            # Set a generic line pattern
            self.line_pattern_input.setText("^(?P<description>.*?)\\s+(?P<quantity>\\d+)\\s+(?P<price>\\d+\\.\\d{2})\\s+(?P<amount>\\d+\\.\\d{2})$")

            # Add line types
            line_types = {
                "quantity": "float",
                "price": "float",
                "amount": "float"
            }

            # Add line types to the table
            for field_name, field_type in line_types.items():
                row_count = self.line_types_table.rowCount()
                self.line_types_table.insertRow(row_count)
                self.line_types_table.setItem(row_count, 0, QTableWidgetItem(field_name))
                type_combo = QComboBox()
                type_combo.addItems(["string", "date", "float", "int"])
                type_combo.setCurrentText(field_type)
                self.line_types_table.setCellWidget(row_count, 1, type_combo)

    def add_keyword(self):
        """Add a keyword to the keywords list"""
        keyword, ok = QInputDialog.getText(self, "Add Keyword", "Enter keyword:")
        if ok and keyword:
            self.keywords_list.addItem(keyword)

    def remove_keyword(self):
        """Remove selected keyword from the list"""
        selected_items = self.keywords_list.selectedItems()
        for item in selected_items:
            self.keywords_list.takeItem(self.keywords_list.row(item))

    def add_exclude_keyword(self):
        """Add an exclude keyword to the list"""
        keyword, ok = QInputDialog.getText(self, "Add Exclude Keyword", "Enter exclude keyword (regex pattern):\nMatching invoices will be skipped")
        if ok and keyword:
            self.exclude_keywords_list.addItem(keyword)

    def remove_exclude_keyword(self):
        """Remove selected exclude keyword from the list"""
        selected_items = self.exclude_keywords_list.selectedItems()
        for item in selected_items:
            self.exclude_keywords_list.takeItem(self.exclude_keywords_list.row(item))

    def add_date_format(self):
        """Add a date format to the list"""
        # Show a dialog with common date formats to choose from
        formats = [
            "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d",
            "%d-%m-%Y", "%m-%d-%Y", "%Y-%m-%d",
            "%d.%m.%Y", "%m.%d.%Y", "%Y.%m.%d",
            "%d %B %Y", "%B %d, %Y", "%Y %B %d",
            "%d %b %Y", "%b %d, %Y", "%Y %b %d"
        ]

        format_str, ok = QInputDialog.getItem(
            self,
            "Add Date Format",
            "Select or enter a date format:\n\n%d = day, %m = month, %Y = year, %B = month name, %b = abbreviated month",
            formats,
            0,
            True  # Allow editing
        )

        if ok and format_str:
            # Check if format already exists
            for i in range(self.date_formats_list.count()):
                if self.date_formats_list.item(i).text() == format_str:
                    return  # Format already exists

            self.date_formats_list.addItem(format_str)

    def remove_date_format(self):
        """Remove selected date format from the list"""
        selected_items = self.date_formats_list.selectedItems()
        for item in selected_items:
            self.date_formats_list.takeItem(self.date_formats_list.row(item))

    def add_table_definition(self):
        """Add a table definition to the tables table"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Table Definition")
        layout = QFormLayout(dialog)

        start_input = QLineEdit()
        end_input = QLineEdit()
        body_input = QLineEdit()

        layout.addRow("Start Pattern:", start_input)
        layout.addRow("End Pattern:", end_input)
        layout.addRow("Body Pattern (with named capture groups):", body_input)

        # Add a help text for named capture groups
        help_text = QLabel("Example body pattern: (?P<hotel_details>[\\S ]+),\\s+(?P<date_check_in>\\d{1,2}\\/\\d{1,2}\\/\\d{4})")
        help_text.setWordWrap(True)
        help_text.setStyleSheet("font-style: italic; color: #666;")
        layout.addRow("", help_text)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)

        if dialog.exec() == QDialog.Accepted:
            start_pattern = start_input.text().strip()
            end_pattern = end_input.text().strip()
            body_pattern = body_input.text().strip()

            if start_pattern and end_pattern and body_pattern:
                row_count = self.tables_definition_table.rowCount()
                self.tables_definition_table.insertRow(row_count)

                # Create QLineEdit widgets for each pattern with live testing
                start_edit = QLineEdit(start_pattern)
                start_edit.textChanged.connect(self.on_regex_pattern_changed)
                self.tables_definition_table.setCellWidget(row_count, 0, start_edit)

                end_edit = QLineEdit(end_pattern)
                end_edit.textChanged.connect(self.on_regex_pattern_changed)
                self.tables_definition_table.setCellWidget(row_count, 1, end_edit)

                body_edit = QLineEdit(body_pattern)
                body_edit.textChanged.connect(self.on_regex_pattern_changed)
                self.tables_definition_table.setCellWidget(row_count, 2, body_edit)

    def remove_table_definition(self):
        """Remove selected table definition from the table"""
        selected_rows = set()
        for item in self.tables_definition_table.selectedItems():
            selected_rows.add(item.row())

        # Remove rows in reverse order to avoid index shifting
        for row in sorted(selected_rows, reverse=True):
            self.tables_definition_table.removeRow(row)

    def add_tax_line_type(self):
        """Add a tax line type to the table"""
        row_count = self.tax_line_types_table.rowCount()
        self.tax_line_types_table.insertRow(row_count)

        # Add field name cell with default values for tax lines
        field_names = ["price_subtotal", "line_tax_percent", "line_tax_amount"]
        if row_count < len(field_names):
            field_name = field_names[row_count]
        else:
            field_name = "field_name"
        self.tax_line_types_table.setItem(row_count, 0, QTableWidgetItem(field_name))

        # Add type cell with combobox
        type_combo = QComboBox()
        type_combo.addItems(["string", "date", "float", "int"])

        # Set default type for tax line fields
        if field_name in ["price_subtotal", "line_tax_percent", "line_tax_amount"]:
            type_combo.setCurrentText("float")

        self.tax_line_types_table.setCellWidget(row_count, 1, type_combo)

    def remove_tax_line_type(self):
        """Remove selected tax line type from the table"""
        selected_rows = set()
        for item in self.tax_line_types_table.selectedItems():
            selected_rows.add(item.row())

        # Remove rows in reverse order to avoid index shifting
        for row in sorted(selected_rows, reverse=True):
            self.tax_line_types_table.removeRow(row)

    def add_field(self):
        """Add a field to the fields table"""
        # Check if we have the required fields already
        required_fields = ['invoice_number', 'date', 'amount']
        existing_fields = []
        for row in range(self.fields_table.rowCount()):
            field_name = self.fields_table.item(row, 0).text().strip()
            existing_fields.append(field_name)

        missing_required = [field for field in required_fields if field not in existing_fields]

        # If we're missing required fields, ask which one to add
        if missing_required:
            field_to_add, ok = QInputDialog.getItem(
                self,
                "Add Required Field",
                "The following fields are required by invoice2data:\n- invoice_number\n- date\n- amount\n\nWhich required field would you like to add?",
                missing_required,
                0,
                False
            )

            if ok and field_to_add:
                row_count = self.fields_table.rowCount()
                self.fields_table.insertRow(row_count)

                # Add field name cell
                self.fields_table.setItem(row_count, 0, QTableWidgetItem(field_to_add))

                # Add regex pattern cell with appropriate default pattern
                if field_to_add == 'invoice_number':
                    pattern = r'(?:Invoice)'
                elif field_to_add == 'date':
                    pattern = r'(?:Date)'
                elif field_to_add == 'amount':
                    pattern = r'(?:Total)'
                else:
                    pattern = 'regex_pattern'

                # Create a QLineEdit for the regex pattern with live testing
                pattern_edit = QLineEdit(pattern)
                pattern_edit.textChanged.connect(self.on_regex_pattern_changed)
                self.fields_table.setCellWidget(row_count, 1, pattern_edit)

                # Add type cell with combobox
                type_combo = QComboBox()
                type_combo.addItems(["string", "date", "float", "int"])

                # Set appropriate type
                if field_to_add == 'invoice_number':
                    type_combo.setCurrentText("string")
                elif field_to_add == 'date':
                    type_combo.setCurrentText("date")
                elif field_to_add == 'amount':
                    type_combo.setCurrentText("float")

                self.fields_table.setCellWidget(row_count, 2, type_combo)

                # If it's a date field, make sure to update the template with formats
                if field_to_add == 'date':
                    self.update_template_preview()

                return

        # Regular field addition if no required fields are missing or user cancelled
        row_count = self.fields_table.rowCount()
        self.fields_table.insertRow(row_count)

        # Add field name cell
        self.fields_table.setItem(row_count, 0, QTableWidgetItem("field_name"))

        # Create a QLineEdit for the regex pattern with live testing
        pattern_edit = QLineEdit("regex_pattern")
        pattern_edit.textChanged.connect(self.on_regex_pattern_changed)
        self.fields_table.setCellWidget(row_count, 1, pattern_edit)

        # Add type cell with combobox
        type_combo = QComboBox()
        type_combo.addItems(["string", "date", "float", "int"])
        self.fields_table.setCellWidget(row_count, 2, type_combo)

    def on_regex_pattern_changed(self, pattern):
        """Handle regex pattern changes in the fields table"""
        # Store the current pattern for highlighting
        self._active_regex_pattern = pattern

        # Highlight matches in the extraction results
        self.highlight_regex_matches(pattern)

    def remove_field(self):
        """Remove selected field from the table"""
        selected_rows = set()
        for item in self.fields_table.selectedItems():
            selected_rows.add(item.row())

        # Remove rows in reverse order to avoid index shifting
        for row in sorted(selected_rows, reverse=True):
            self.fields_table.removeRow(row)

    def add_firstline_pattern(self):
        """Add a first line pattern to the list"""
        pattern, ok = QInputDialog.getText(self, "Add First Line Pattern", "Enter regex pattern:")
        if ok and pattern:
            self.firstline_list.addItem(pattern)

    def remove_firstline_pattern(self):
        """Remove selected first line pattern from the list"""
        selected_items = self.firstline_list.selectedItems()
        for item in selected_items:
            self.firstline_list.takeItem(self.firstline_list.row(item))

    def add_line_type(self):
        """Add a line type to the table"""
        row_count = self.line_types_table.rowCount()
        self.line_types_table.insertRow(row_count)

        # Add field name cell
        self.line_types_table.setItem(row_count, 0, QTableWidgetItem("field_name"))

        # Add type cell with combobox
        type_combo = QComboBox()
        type_combo.addItems(["string", "date", "float", "int"])
        self.line_types_table.setCellWidget(row_count, 1, type_combo)

    def remove_line_type(self):
        """Remove selected line type from the table"""
        selected_rows = set()
        for item in self.line_types_table.selectedItems():
            selected_rows.add(item.row())

        # Remove rows in reverse order to avoid index shifting
        for row in sorted(selected_rows, reverse=True):
            self.line_types_table.removeRow(row)

    def update_template_preview(self):
        """Update the template preview based on current inputs"""
        try:
            # Build the template dictionary
            template = self.build_invoice2data_template()

            # Print debug information
            print(f"[DEBUG] Built template with keys: {list(template.keys())}")
            print(f"[DEBUG] Template issuer: {template.get('issuer', 'None')}")
            print(f"[DEBUG] Template fields: {list(template.get('fields', {}).keys())}")
            print(f"[DEBUG] Template keywords: {template.get('keywords', [])}")

            # Convert to YAML for preview (preferred format)
            import yaml

            # Make a deep copy of the template to avoid modifying the original
            import copy
            template_copy = copy.deepcopy(template)

            # Clean up empty values to improve YAML readability
            # Remove empty lists and dictionaries
            for key in list(template_copy.keys()):
                if isinstance(template_copy[key], list) and not template_copy[key]:
                    template_copy[key] = None
                elif isinstance(template_copy[key], dict):
                    # Clean up nested dictionaries
                    for nested_key in list(template_copy[key].keys()):
                        if template_copy[key][nested_key] == "" or template_copy[key][nested_key] == [] or template_copy[key][nested_key] == {}:
                            del template_copy[key][nested_key]
                    # Remove empty dictionaries
                    if not template_copy[key]:
                        template_copy[key] = None

            # Use safe_dump with explicit indentation for better readability
            template_yaml = yaml.safe_dump(
                template_copy,
                default_flow_style=False,
                allow_unicode=True,
                sort_keys=False,  # Preserve key order
                indent=2,         # Explicit indentation
                width=80          # Line width
            )

            print(f"[DEBUG] Generated YAML preview (first 100 chars): {template_yaml[:100]}...")

            # Set the YAML text in the preview
            self.template_preview.clear()  # Clear existing content first
            self.template_preview.setText(template_yaml)

            # Force update of the preview widget
            self.template_preview.repaint()

            # Ensure the preview is visible
            if hasattr(self, 'template_tab_widget'):
                # Find the template tab index
                for i in range(self.template_tab_widget.count()):
                    if self.template_tab_widget.tabText(i) == "Template":
                        self.template_tab_widget.setCurrentIndex(i)
                        break

            # Store the template for later use (e.g., when testing)
            self._last_valid_template = template

            print(f"[DEBUG] Template preview updated successfully")

        except Exception as e:
            import traceback
            error_message = f"Error generating preview: {str(e)}\n\n{traceback.format_exc()}"
            print(error_message)
            self.template_preview.setText(f"Error generating preview: {str(e)}")

    def apply_template_changes(self):
        """Apply changes from the edited template preview to the form"""
        try:
            # Get the edited YAML from the preview
            template_yaml = self.template_preview.toPlainText()

            # Try to parse as YAML first (preferred format)
            try:
                import yaml
                # Use safe_load to handle potential security issues
                template = yaml.safe_load(template_yaml)
                print(f"[DEBUG] Successfully parsed template as YAML: {type(template)}")

                # Print the first few keys for debugging
                if isinstance(template, dict):
                    print(f"[DEBUG] Template keys: {list(template.keys())[:5]}")
                else:
                    print(f"[DEBUG] Template is not a dictionary: {type(template)}")
                    raise ValueError("Template must be a dictionary")

            except yaml.YAMLError as e:
                print(f"[DEBUG] YAML parsing error: {str(e)}")
                # If YAML parsing fails, try JSON as fallback
                try:
                    import json
                    template = json.loads(template_yaml)
                    print(f"[DEBUG] Successfully parsed template as JSON")
                except json.JSONDecodeError as json_e:
                    print(f"[DEBUG] JSON parsing error: {str(json_e)}")
                    QMessageBox.critical(
                        self,
                        "Invalid Format",
                        f"The template contains invalid YAML and JSON and cannot be applied.\n\nYAML Error: {str(e)}\n\nJSON Error: {str(json_e)}"
                    )
                    return

            # Check if template is None
            if template is None:
                QMessageBox.critical(
                    self,
                    "Invalid Template",
                    "The template could not be parsed. Please check the format and try again."
                )
                return

            # Validate required fields for invoice2data template
            required_fields = ["issuer", "fields", "keywords"]
            missing_fields = [field for field in required_fields if field not in template]
            if missing_fields:
                QMessageBox.warning(
                    self,
                    "Missing Required Fields",
                    f"The template is missing the following required fields: {', '.join(missing_fields)}.\n\nThese fields will be added with default values."
                )
                # Add missing required fields with default values
                if "issuer" not in template:
                    template["issuer"] = "Unknown Vendor"
                if "fields" not in template:
                    template["fields"] = {}
                if "keywords" not in template:
                    template["keywords"] = []

            # Clean up empty values to improve YAML readability
            for key in list(template.keys()):
                if isinstance(template[key], list) and not template[key]:
                    template[key] = None
                elif isinstance(template[key], dict):
                    # Clean up nested dictionaries
                    for nested_key in list(template[key].keys()):
                        if template[key][nested_key] == "" or template[key][nested_key] == [] or template[key][nested_key] == {}:
                            del template[key][nested_key]
                    # Remove empty dictionaries
                    if not template[key]:
                        template[key] = None

            # Format the template as YAML and update the preview
            formatted_yaml = yaml.safe_dump(
                template,
                default_flow_style=False,
                allow_unicode=True,
                sort_keys=False,  # Preserve key order
                indent=2,         # Explicit indentation
                width=80          # Line width
            )

            # Update the preview with the formatted YAML
            self.template_preview.setText(formatted_yaml)
            print(f"[DEBUG] Updated template preview with formatted YAML")

            # Store the successfully parsed and formatted template
            self._last_valid_template = template

            # Update the form fields based on the template
            # General tab
            if "issuer" in template:
                self.issuer_input.setText(template.get("issuer", ""))

            # Set priority if available
            if "priority" in template:
                try:
                    priority_value = int(template.get("priority", 100))
                    self.priority_input.setValue(priority_value)
                except (ValueError, TypeError):
                    print(f"[WARNING] Invalid priority value: {template.get('priority')}, using default 100")
                    self.priority_input.setValue(100)

            # Keywords
            if "keywords" in template:
                self.keywords_list.clear()
                keywords = template.get("keywords", [])
                # Handle None value
                if keywords is None:
                    print("[WARNING] keywords is None, treating as empty list")
                    keywords = []
                for keyword in keywords:
                    self.keywords_list.addItem(keyword)

            # Exclude Keywords
            if "exclude_keywords" in template:
                self.exclude_keywords_list.clear()
                exclude_keywords = template.get("exclude_keywords", [])
                # Handle None value
                if exclude_keywords is None:
                    print("[WARNING] exclude_keywords is None, treating as empty list")
                    exclude_keywords = []
                for keyword in exclude_keywords:
                    self.exclude_keywords_list.addItem(keyword)

            # Options
            if "options" in template:
                options = template.get("options", {})
                # Handle None value
                if options is None:
                    print("[WARNING] options is None, treating as empty dict")
                    options = {}

                self.currency_input.setText(options.get("currency", "INR"))
                self.decimal_separator_input.setText(options.get("decimal_separator", "."))

                # Set language if available
                languages = options.get("languages")
                if languages is not None and len(languages) > 0:
                    language = languages[0]
                    index = self.language_input.findText(language)
                    if index >= 0:
                        self.language_input.setCurrentIndex(index)

                # Set replace patterns if available
                replace_patterns = options.get("replace")
                if replace_patterns is not None and replace_patterns:
                    # Clear existing patterns
                    while self.replace_table.rowCount() > 0:
                        self.replace_table.removeRow(0)

                    # Add patterns from template
                    for pattern in replace_patterns:
                        if isinstance(pattern, list) and len(pattern) >= 2:
                            row = self.replace_table.rowCount()
                            self.replace_table.insertRow(row)
                            self.replace_table.setItem(row, 0, QTableWidgetItem(pattern[0]))
                            self.replace_table.setItem(row, 1, QTableWidgetItem(pattern[1]))

                # Set date formats if available
                date_formats = options.get("date_formats")
                if date_formats is not None and date_formats:
                    # Clear existing date formats
                    self.date_formats_list.clear()

                    # Add date formats from template
                    for format_str in date_formats:
                        self.date_formats_list.addItem(format_str)

                # Set additional options
                self.remove_whitespace_checkbox.setChecked(options.get("remove_whitespace", False))
                self.remove_accents_checkbox.setChecked(options.get("remove_accents", False))
                self.lowercase_checkbox.setChecked(options.get("lowercase", False))

            # Lines
            if "lines" in template:
                lines = template.get("lines", {})
                # Handle None value
                if lines is None:
                    print("[WARNING] lines is None, treating as empty dict")
                    lines = {}

                # Start, end, line patterns
                if "start" in lines:
                    self.line_start_input.setText(lines.get("start", ""))
                else:
                    self.line_start_input.clear()

                if "end" in lines:
                    self.line_end_input.setText(lines.get("end", ""))
                else:
                    self.line_end_input.clear()

                if "line" in lines:
                    self.line_pattern_input.setText(lines.get("line", ""))
                else:
                    self.line_pattern_input.clear()

                if "skip_line" in lines:
                    self.skip_line_input.setText(lines.get("skip_line", ""))
                else:
                    self.skip_line_input.clear()

                # First line patterns
                self.firstline_list.clear()
                first_lines = lines.get("first_line", [])
                # Handle None value
                if first_lines is None:
                    print("[WARNING] first_line is None, treating as empty list")
                    first_lines = []
                for first_line in first_lines:
                    self.firstline_list.addItem(first_line)

                # Line types
                # Clear existing line types
                while self.line_types_table.rowCount() > 0:
                    self.line_types_table.removeRow(0)

                # Add line types from template
                line_types = lines.get("types", {})
                # Handle None value
                if line_types is None:
                    print("[WARNING] line types is None, treating as empty dict")
                    line_types = {}
                for field_name, field_type in line_types.items():
                    row = self.line_types_table.rowCount()
                    self.line_types_table.insertRow(row)
                    self.line_types_table.setItem(row, 0, QTableWidgetItem(field_name))

                    # Add type combo box
                    type_combo = QComboBox()
                    type_combo.addItems(["string", "date", "float", "int"])
                    type_combo.setCurrentText(field_type)
                    self.line_types_table.setCellWidget(row, 1, type_combo)

            # Tables
            if "tables" in template:
                # Clear existing table definitions
                while self.tables_definition_table.rowCount() > 0:
                    self.tables_definition_table.removeRow(0)

                # Add table definitions from template
                tables = template.get("tables", [])
                # Handle None value
                if tables is None:
                    print("[WARNING] tables is None, treating as empty list")
                    tables = []
                for table_def in tables:
                    if isinstance(table_def, dict) and "start" in table_def and "end" in table_def and "body" in table_def:
                        row = self.tables_definition_table.rowCount()
                        self.tables_definition_table.insertRow(row)
                        self.tables_definition_table.setItem(row, 0, QTableWidgetItem(table_def.get("start", "")))
                        self.tables_definition_table.setItem(row, 1, QTableWidgetItem(table_def.get("end", "")))
                        self.tables_definition_table.setItem(row, 2, QTableWidgetItem(table_def.get("body", "")))

            # Tax Lines
            if "tax_lines" in template:
                tax_lines = template.get("tax_lines", {})
                # Handle None value
                if tax_lines is None:
                    print("[WARNING] tax_lines is None, treating as empty dict")
                    tax_lines = {}

                if "start" in tax_lines:
                    self.tax_lines_start_input.setText(tax_lines.get("start", ""))
                else:
                    self.tax_lines_start_input.clear()

                if "end" in tax_lines:
                    self.tax_lines_end_input.setText(tax_lines.get("end", ""))
                else:
                    self.tax_lines_end_input.clear()

                if "line" in tax_lines:
                    self.tax_lines_line_input.setText(tax_lines.get("line", ""))
                else:
                    self.tax_lines_line_input.clear()

                # Tax line types
                # Clear existing tax line types
                while self.tax_line_types_table.rowCount() > 0:
                    self.tax_line_types_table.removeRow(0)

                # Add tax line types from template
                if "types" in tax_lines:
                    tax_types = tax_lines.get("types", {})
                    # Handle None value
                    if tax_types is None:
                        print("[WARNING] tax line types is None, treating as empty dict")
                        tax_types = {}
                    for field_name, field_type in tax_types.items():
                        row = self.tax_line_types_table.rowCount()
                        self.tax_line_types_table.insertRow(row)
                        self.tax_line_types_table.setItem(row, 0, QTableWidgetItem(field_name))

                        # Add type combo box
                        type_combo = QComboBox()
                        type_combo.addItems(["string", "date", "float", "int"])
                        type_combo.setCurrentText(field_type)
                        self.tax_line_types_table.setCellWidget(row, 1, type_combo)

            # Fields
            if "fields" in template:
                # Clear existing fields
                while self.fields_table.rowCount() > 0:
                    self.fields_table.removeRow(0)

                # Add fields from template
                fields = template.get("fields", {})
                # Handle None value
                if fields is None:
                    print("[WARNING] fields is None, treating as empty dict")
                    fields = {}
                for field_name, field_data in fields.items():
                    row = self.fields_table.rowCount()
                    self.fields_table.insertRow(row)
                    self.fields_table.setItem(row, 0, QTableWidgetItem(field_name))

                    # Handle different field data formats
                    if isinstance(field_data, dict):
                        # Complex field with regex and type
                        regex = field_data.get("regex", "")
                        field_type = field_data.get("type", "string")
                    else:
                        # Simple field (just a value)
                        regex = str(field_data)
                        field_type = "string"

                    self.fields_table.setItem(row, 1, QTableWidgetItem(regex))

                    # Add type combo box
                    type_combo = QComboBox()
                    type_combo.addItems(["string", "date", "float", "int"])
                    type_combo.setCurrentText(field_type)
                    self.fields_table.setCellWidget(row, 2, type_combo)

            # Show success message
            QMessageBox.information(
                self,
                "Changes Applied",
                "The template changes have been applied to the form."
            )

        except Exception as e:
            print(f"[ERROR] Failed to apply template changes: {str(e)}")
            import traceback
            traceback.print_exc()

            QMessageBox.critical(
                self,
                "Error",
                f"Failed to apply template changes: {str(e)}"
            )

    def add_replace_pattern(self):
        """Add a replace pattern to the table"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Replace Pattern")
        layout = QFormLayout(dialog)

        find_input = QLineEdit()
        replace_input = QLineEdit()

        layout.addRow("Find (regex):", find_input)
        layout.addRow("Replace with:", replace_input)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)

        if dialog.exec() == QDialog.Accepted:
            find_pattern = find_input.text().strip()
            replace_pattern = replace_input.text().strip()
            if find_pattern:
                row_count = self.replace_table.rowCount()
                self.replace_table.insertRow(row_count)
                self.replace_table.setItem(row_count, 0, QTableWidgetItem(find_pattern))
                self.replace_table.setItem(row_count, 1, QTableWidgetItem(replace_pattern))

    def remove_replace_pattern(self):
        """Remove selected replace pattern from the table"""
        selected_rows = set()
        for item in self.replace_table.selectedItems():
            selected_rows.add(item.row())

        # Remove rows in reverse order to avoid index shifting
        for row in sorted(selected_rows, reverse=True):
            self.replace_table.removeRow(row)

    def build_invoice2data_template(self):
        """Build the invoice2data template dictionary from UI inputs"""
        # Get issuer name (used for template name too)
        issuer = self.issuer_input.text().strip()

        # Get priority value
        priority = self.priority_input.value()

        # Build fields data
        fields_data = {}
        for row in range(self.fields_table.rowCount()):
            field_name = self.fields_table.item(row, 0).text().strip()

            # Get regex pattern - handle both QLineEdit and QTableWidgetItem
            regex_pattern_widget = self.fields_table.cellWidget(row, 1)
            if regex_pattern_widget and isinstance(regex_pattern_widget, QLineEdit):
                regex_pattern = regex_pattern_widget.text().strip()
            else:
                regex_pattern = self.fields_table.item(row, 1).text().strip()

            field_type = self.fields_table.cellWidget(row, 2).currentText()

            # Validate field type - ensure it's one of the supported types
            valid_types = ["string", "date", "float", "int"]
            if field_type not in valid_types:
                print(f"[DEBUG] Invalid field type '{field_type}' for field '{field_name}', defaulting to 'string'")
                field_type = "string"

            # Special handling for required fields
            if field_name == "invoice_number":
                # Use simple string pattern for invoice_number
                if not regex_pattern or regex_pattern == "regex_pattern":
                    regex_pattern = r'Invoice\s+Number\s*:?\s*([\w\-\/]+)'
                fields_data[field_name] = regex_pattern
                print(f"[DEBUG] Using pattern for invoice_number: {regex_pattern}")

            elif field_name == "date":
                # Use complex object with type for date field
                if not regex_pattern or regex_pattern == "regex_pattern":
                    regex_pattern = r'Date\s*:?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})'

                # Get date formats
                date_formats = []
                if self.date_formats_list.count() > 0:
                    for i in range(self.date_formats_list.count()):
                        format_text = self.date_formats_list.item(i).text().strip()
                        if format_text:  # Only add non-empty formats
                            date_formats.append(format_text)

                fields_data[field_name] = {
                    "parser": "regex",
                    "regex": regex_pattern,
                    "type": "date",
                    "formats": date_formats
                }
                print(f"[DEBUG] Using pattern for date: {regex_pattern} with formats: {date_formats}")

            elif field_name == "amount":
                # Use complex object with type for amount field
                if not regex_pattern or regex_pattern == "regex_pattern":
                    regex_pattern = r'GRAND\s+TOTAL\s*[\$\€\£]?\s*(\d+[\.,]\d+)'
                fields_data[field_name] = {
                    "parser": "regex",
                    "regex": regex_pattern,
                    "type": "float"
                }
                print(f"[DEBUG] Using pattern for amount: {regex_pattern}")

            elif field_type == "string":
                # Use simple string pattern for string fields
                fields_data[field_name] = regex_pattern
                print(f"[DEBUG] Using simple string pattern for field '{field_name}'")
            else:
                # Use complex object with type for non-string fields
                field_obj = {
                    "parser": "regex",
                    "regex": regex_pattern,
                    "type": field_type
                }

                # Add formats for date fields
                if field_type == "date":
                    date_formats = []
                    if self.date_formats_list.count() > 0:
                        for i in range(self.date_formats_list.count()):
                            format_text = self.date_formats_list.item(i).text().strip()
                            if format_text:  # Only add non-empty formats
                                date_formats.append(format_text)
                    if date_formats:
                        field_obj["formats"] = date_formats

                fields_data[field_name] = field_obj
                print(f"[DEBUG] Using complex object with type '{field_type}' for field '{field_name}'")

        # Build options data
        options_data = {}

        # Add currency
        currency = self.currency_input.text().strip()
        if currency:
            options_data["currency"] = currency

        # Add language
        language = self.language_input.currentText()
        if language:
            options_data["languages"] = [language]

        # Add decimal separator
        decimal_separator = self.decimal_separator_input.text().strip()
        if decimal_separator and decimal_separator != ".":
            options_data["decimal_separator"] = decimal_separator

        # Add checkbox options
        if self.remove_whitespace_checkbox.isChecked():
            options_data["remove_whitespace"] = True
        if self.remove_accents_checkbox.isChecked():
            options_data["remove_accents"] = True
        if self.lowercase_checkbox.isChecked():
            options_data["lowercase"] = True

        # Add replace patterns
        replace_patterns = []
        for row in range(self.replace_table.rowCount()):
            find_pattern = self.replace_table.item(row, 0).text().strip()
            replace_pattern = self.replace_table.item(row, 1).text().strip()
            if find_pattern:  # Only add if find pattern is not empty
                replace_patterns.append([find_pattern, replace_pattern])
        if replace_patterns:
            options_data["replace"] = replace_patterns

        # Add date formats
        date_formats = []
        for i in range(self.date_formats_list.count()):
            format_text = self.date_formats_list.item(i).text().strip()
            if format_text:  # Only add non-empty formats
                date_formats.append(format_text)
        if date_formats:
            options_data["date_formats"] = date_formats

        # Build keywords list
        keywords = []
        # Always include issuer as a keyword if not already in the list
        issuer_found = False

        for i in range(self.keywords_list.count()):
            keyword = self.keywords_list.item(i).text().strip()
            if keyword:  # Only add non-empty keywords
                keywords.append(keyword)
                if keyword.lower() == issuer.lower():
                    issuer_found = True

        # Add issuer if not already in the list
        if not issuer_found and issuer:
            keywords.append(issuer)

        # Build exclude_keywords list
        exclude_keywords = []
        for i in range(self.exclude_keywords_list.count()):
            keyword = self.exclude_keywords_list.item(i).text().strip()
            if keyword:  # Only add non-empty keywords
                exclude_keywords.append(keyword)

        # Build lines data
        lines_data = {}

        # Check for non-empty start, end, line, and skip_line patterns
        start_line = self.line_start_input.text().strip()
        end_line = self.line_end_input.text().strip()
        line_pattern = self.line_pattern_input.text().strip()
        skip_line = self.skip_line_input.text().strip()

        # Add lines section if any content is defined
        if start_line or end_line or line_pattern or skip_line or self.firstline_list.count() > 0 or self.line_types_table.rowCount() > 0:
            lines_data["parser"] = "lines"

            # Add start, end, line, and skip_line patterns if they exist
            if start_line:
                lines_data["start"] = start_line
            if end_line:
                lines_data["end"] = end_line
            if line_pattern:
                lines_data["line"] = line_pattern
            if skip_line:
                lines_data["skip_line"] = skip_line

            # Add first line patterns if any exist
            if self.firstline_list.count() > 0:
                lines_data["first_line"] = []
                for i in range(self.firstline_list.count()):
                    pattern = self.firstline_list.item(i).text().strip()
                    if pattern:  # Only add non-empty patterns
                        lines_data["first_line"].append(pattern)

            # Add line types if any exist
            if self.line_types_table.rowCount() > 0:
                lines_data["types"] = {}
                for row in range(self.line_types_table.rowCount()):
                    field_name = self.line_types_table.item(row, 0).text().strip()
                    field_type = self.line_types_table.cellWidget(row, 1).currentText()
                    if field_name:
                        lines_data["types"][field_name] = field_type

        # Add tables if any exist
        if self.tables_definition_table.rowCount() > 0:
            tables_data = []
            for row in range(self.tables_definition_table.rowCount()):
                # Get start pattern - handle both QLineEdit and QTableWidgetItem
                start_pattern_widget = self.tables_definition_table.cellWidget(row, 0)
                if start_pattern_widget and isinstance(start_pattern_widget, QLineEdit):
                    start_pattern = start_pattern_widget.text().strip()
                else:
                    start_pattern = self.tables_definition_table.item(row, 0).text().strip()

                # Get end pattern - handle both QLineEdit and QTableWidgetItem
                end_pattern_widget = self.tables_definition_table.cellWidget(row, 1)
                if end_pattern_widget and isinstance(end_pattern_widget, QLineEdit):
                    end_pattern = end_pattern_widget.text().strip()
                else:
                    end_pattern = self.tables_definition_table.item(row, 1).text().strip()

                # Get body pattern - handle both QLineEdit and QTableWidgetItem
                body_pattern_widget = self.tables_definition_table.cellWidget(row, 2)
                if body_pattern_widget and isinstance(body_pattern_widget, QLineEdit):
                    body_pattern = body_pattern_widget.text().strip()
                else:
                    body_pattern = self.tables_definition_table.item(row, 2).text().strip()

                if start_pattern and end_pattern and body_pattern:
                    tables_data.append({
                        "start": start_pattern,
                        "end": end_pattern,
                        "body": body_pattern
                    })
            if tables_data:
                lines_data["tables"] = tables_data

        # Add tax lines if any exist
        tax_lines_start = self.tax_lines_start_input.text().strip()
        tax_lines_end = self.tax_lines_end_input.text().strip()
        tax_lines_line = self.tax_lines_line_input.text().strip()

        if tax_lines_start or tax_lines_end or tax_lines_line or self.tax_line_types_table.rowCount() > 0:
            tax_lines_data = {
                "parser": "lines"
            }

            # Add start, end, and line patterns if they exist
            if tax_lines_start:
                tax_lines_data["start"] = tax_lines_start
            if tax_lines_end:
                tax_lines_data["end"] = tax_lines_end
            if tax_lines_line:
                tax_lines_data["line"] = tax_lines_line

            # Add tax line types if any exist
            if self.tax_line_types_table.rowCount() > 0:
                tax_lines_data["types"] = {}
                for row in range(self.tax_line_types_table.rowCount()):
                    field_name = self.tax_line_types_table.item(row, 0).text().strip()
                    field_type = self.tax_line_types_table.cellWidget(row, 1).currentText()
                    if field_name:
                        tax_lines_data["types"][field_name] = field_type

            lines_data["tax_lines"] = tax_lines_data

        # Use the unified invoice processing utilities to build the template
        template = invoice_processing_utils.build_invoice2data_template(
            issuer=issuer,
            fields_data=fields_data,
            options_data=options_data,
            keywords=keywords,
            exclude_keywords=exclude_keywords if exclude_keywords else None,
            lines_data=lines_data if lines_data else None
        )

        # Add priority to the template
        template["priority"] = priority

        return template

    def test_invoice2data_template(self):
        """Test the current template against the loaded PDF using invoice2data API directly"""
        # Import required modules
        import re
        import json
        import yaml

        if not self.pdf_path:
            QMessageBox.warning(self, "Warning", "Please load a PDF first")
            return

        # Process events to prevent UI freezing
        print("[DEBUG] Testing invoice2data template...")
        QApplication.processEvents()

        try:
            # Get the template from the preview
            template_text = self.template_preview.toPlainText()

            try:
                # Try to parse as YAML first (preferred format)
                import yaml
                template = yaml.safe_load(template_text)
                print("[DEBUG] Successfully parsed template from preview as YAML")
            except yaml.YAMLError as e:
                # If YAML parsing fails, try JSON as fallback
                try:
                    import json
                    template = json.loads(template_text)
                    print("[DEBUG] Successfully parsed template from preview as JSON")
                except json.JSONDecodeError as json_e:
                    # Show error message if both YAML and JSON parsing fail
                    QMessageBox.critical(
                        self,
                        "Invalid Template Format",
                        f"The template preview contains invalid YAML/JSON and cannot be used.\n\nYAML Error: {str(e)}\n\nJSON Error: {str(json_e)}"
                    )
                    return

            # Create a temporary directory for our files using path_helper
            if PATH_HELPER_AVAILABLE:
                temp_dir = path_helper.ensure_directory("templates")
            else:
                temp_dir = os.path.abspath("templates")
                os.makedirs(temp_dir, exist_ok=True)

            # Clean up any existing files in the template directory
            print("[DEBUG] Cleaning template directory before creating new files")
            for file in os.listdir(temp_dir):
                if file.startswith("test") and (file.endswith(".json") or file.endswith(".yml") or file.endswith(".txt")):
                    os.remove(os.path.join(temp_dir, file))
                    print(f"[DEBUG] Removed old file: {os.path.join(temp_dir, file)}")

            # Create a unique base filename for our test
            base_filename = "test_template"

            # Create temporary template file
            temp_template_path = os.path.join(temp_dir, f"{base_filename}.yml")
            print(f"[DEBUG] Using template path: {temp_template_path}")

            # Make sure the template is in the correct format
            if isinstance(template, str):
                try:
                    # Try YAML first
                    import yaml
                    template = yaml.safe_load(template)
                    print("[DEBUG] Converted template from string to object using YAML")
                except yaml.YAMLError:
                    try:
                        # Fallback to JSON
                        template = json.loads(template)
                        print("[DEBUG] Converted template from string to object using JSON")
                    except json.JSONDecodeError as e:
                        print(f"[DEBUG] Error parsing template string: {str(e)}")

            # Create a new template dictionary with the correct structure for invoice2data
            new_template = {}

            # Copy fields from the original template
            if "issuer" in template:
                new_template["issuer"] = template["issuer"]
            else:
                new_template["issuer"] = "Test Template"

            # Set required fields
            new_template["name"] = base_filename
            new_template["template_name"] = base_filename

            # Keywords are required - use existing keywords if available, otherwise use issuer and filename
            if "keywords" in template and template["keywords"]:
                # Use existing keywords
                new_template["keywords"] = template["keywords"]
            else:
                # Use issuer and filename as keywords
                keywords = []
                if "issuer" in new_template and new_template["issuer"]:
                    keywords.append(new_template["issuer"])
                keywords.append(base_filename)
                new_template["keywords"] = keywords

            # Add exclude_keywords if missing (required by InvoiceTemplate)
            if "exclude_keywords" in template:
                new_template["exclude_keywords"] = template["exclude_keywords"]
            else:
                new_template["exclude_keywords"] = []

            # Copy fields
            if "fields" in template:
                new_template["fields"] = template["fields"]
            else:
                new_template["fields"] = {}

            # Copy options
            if "options" in template:
                new_template["options"] = template["options"]
            else:
                new_template["options"] = {"currency": "USD", "languages": ["en"]}

            # Copy other fields
            if "lines" in template:
                new_template["lines"] = template["lines"]
            if "priority" in template:
                new_template["priority"] = template["priority"]

            # Replace the original template with the new one
            template = new_template

            # Save as YAML with improved formatting
            try:
                with open(temp_template_path, "w", encoding="utf-8") as f:
                    yaml.safe_dump(
                        template,
                        f,
                        default_flow_style=False,
                        allow_unicode=True,
                        sort_keys=False,  # Preserve key order
                        indent=2,         # Explicit indentation
                        width=80          # Line width
                    )
                print(f"[DEBUG] Saved template to {temp_template_path}")

                # Verify the template is valid YAML
                with open(temp_template_path, "r", encoding="utf-8") as f:
                    yaml_content = yaml.safe_load(f)
                print(f"[DEBUG] Verified template is valid YAML with keys: {list(yaml_content.keys())}")

                # Print the template content for debugging
                with open(temp_template_path, "r", encoding="utf-8") as f:
                    template_content = f.read()
                print(f"[DEBUG] Template content (first 200 chars):\n{template_content[:200]}...")
            except Exception as yaml_error:
                print(f"[DEBUG] Error saving YAML template: {str(yaml_error)}")
                self.template_preview.setText(f"TEST RESULT - ERROR\n\nError saving template: {str(yaml_error)}")
                return

            # Ensure we're using the latest cached extraction data
            # This is especially important after using the "Adjust Parameters" button
            # which can completely change the extracted data
            print("[DEBUG] Using latest cached extraction data for invoice2data")

            # Get the latest cached extraction data directly
            # This ensures we use the most recent data without forcing a new extraction
            current_data = self._get_current_json_data()

            # Add metadata to the data
            if self.pdf_path and current_data:
                current_data['metadata'] = {
                    'filename': os.path.basename(self.pdf_path),
                    'page_count': len(self.pdf_document) if self.pdf_document else 1,
                    'template_type': 'multi' if self.multi_page_mode else 'single',
                    'creation_date': datetime.datetime.now().isoformat()
                }

            # Log the data structure for debugging
            print(f"[DEBUG] Current data keys: {list(current_data.keys())}")
            if 'header' in current_data:
                print(f"[DEBUG] Header data type: {type(current_data['header'])}")
            if 'items' in current_data:
                print(f"[DEBUG] Items data type: {type(current_data['items'])}")
            if 'summary' in current_data:
                print(f"[DEBUG] Summary data type: {type(current_data['summary'])}")

            # Convert the latest data to text format for invoice2data using the unified utilities
            pdf_text = invoice_processing_utils.convert_extraction_to_text(current_data, pdf_path=self.pdf_path)

            # Save the text to a temporary file
            temp_text_path = os.path.join(temp_dir, f"{base_filename}.txt")
            with open(temp_text_path, "w", encoding="utf-8") as f:
                f.write(pdf_text)
            print(f"[DEBUG] Saved extracted text to {temp_text_path}")

            # Use the unified invoice processing utilities
            try:
                # Create template data dictionary in the format expected by process_with_invoice2data
                template_data = {"json_template": template}

                # Process with invoice2data using the unified utilities
                print(f"[DEBUG] Using invoice_processing_utils.process_with_invoice2data")
                result = invoice_processing_utils.process_with_invoice2data(
                    pdf_path=self.pdf_path,
                    template_data=template_data,
                    extracted_data=current_data,
                    temp_dir=temp_dir
                )

                print(f"[DEBUG] Extraction result: {result}")

                # If the shared method returns None, we'll consider it a failure
                if result is None:
                    print(f"[DEBUG] Shared method returned None, extraction failed")
                    result = None

                if result:
                    # Convert result to JSON for display
                    result_json = json.dumps(result, indent=2, default=str)

                    # Display the result in the template preview
                    self.template_preview.setText(f"TEST RESULT - SUCCESS\n\n{result_json}")

                    # Show success message
                    QMessageBox.information(
                        self,
                        "Test Successful",
                        "The template successfully extracted data from the invoice."
                    )
                else:
                    # No result - show detailed error with suggestions

                    # Get the warnings from invoice2data
                    import io
                    import logging
                    invoice2data_logger = logging.getLogger("invoice2data")
                    log_capture = io.StringIO()
                    for handler in invoice2data_logger.handlers:
                        if isinstance(handler, logging.StreamHandler) and handler.stream != sys.stderr and handler.stream != sys.stdout:
                            log_content = handler.stream.getvalue()
                            if log_content:
                                log_capture.write(log_content)

                    warnings_text = log_capture.getvalue()
                    print(f"[DEBUG] Captured warnings: {warnings_text}")

                    # Analyze the warnings and provide suggestions
                    suggestions = self.analyze_invoice2data_warnings(warnings_text, template)

                    # Check the text file content for debugging
                    text_content = ""
                    try:
                        with open(temp_text_path, 'r', encoding='utf-8') as f:
                            text_content = f.read()
                    except Exception as e:
                        print(f"[DEBUG] Error reading text file: {str(e)}")

                    # Create a detailed error message with suggestions
                    error_message = "TEST RESULT - ERROR\n\n"
                    error_message += "No data was extracted. Here are some suggestions to fix your template:\n\n"

                    # Add specific suggestions based on the warnings
                    error_message += suggestions + "\n\n"

                    # Add general suggestions
                    error_message += "General suggestions:\n"
                    error_message += "1. Make sure your keywords match text in the invoice\n"
                    error_message += "2. Check that your regex patterns have capturing groups (parentheses)\n"
                    error_message += "3. Verify that the field names match those in the invoice\n"
                    error_message += "4. For date fields, ensure the formats list includes the format used in the invoice\n\n"

                    # Add a sample of the text content for reference
                    if text_content:
                        error_message += "Invoice text sample (first 500 chars):\n"
                        error_message += text_content[:500] + "...\n\n"

                    # Display the detailed error message in the template preview
                    self.template_preview.setText(error_message)

                    # Show error message with suggestions
                    QMessageBox.warning(
                        self,
                        "Test Error",
                        "No data was extracted. See the template preview for suggestions on how to fix your template."
                    )
            except Exception as api_error:
                print(f"[DEBUG] Error using invoice2data API: {str(api_error)}")
                import traceback
                traceback_text = traceback.format_exc()
                print(traceback_text)

                # Try using the command-line approach as a fallback
                print("[DEBUG] Trying command-line approach as fallback...")
                try:
                    import subprocess

                    # Build the command
                    template_dir = temp_dir if 'temp_dir' in locals() else os.path.dirname(temp_template_path)
                    cmd = [
                        "invoice2data",
                        "--input-reader", "text",
                        "--exclude-built-in-templates",
                        "--template-folder", template_dir,
                        "--output-format", "json",
                        temp_text_path
                    ]
                    print(f"[DEBUG] Using template directory for command line: {template_dir}")
                    print(f"[DEBUG] Running command: {' '.join(cmd)}")

                    # Run the command
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        check=False
                    )

                    # Check if command was successful
                    if result.returncode == 0:
                        print(f"[DEBUG] Command succeeded with output: {result.stdout}")

                        # Try to find JSON in the output
                        json_match = re.search(r'(\{.*\})', result.stdout, re.DOTALL)
                        if json_match:
                            json_str = json_match.group(1)
                            extraction_result = json.loads(json_str)
                            result_json = json.dumps(extraction_result, indent=2)

                            # Display the result
                            self.template_preview.setText(f"TEST RESULT - SUCCESS (via command line)\n\n{result_json}")

                            # Show success message
                            QMessageBox.information(
                                self,
                                "Test Successful",
                                "The template successfully extracted data from the invoice (using command-line fallback)."
                            )
                            return

                    # Command-line approach failed
                    print(f"[DEBUG] Command-line approach failed: {result.stderr}")

                    # Get the warnings from stderr
                    warnings_text = result.stderr

                    # Analyze the warnings and provide suggestions
                    suggestions = self.analyze_invoice2data_warnings(warnings_text, template)

                    # Check the text file content for debugging
                    text_content = ""
                    try:
                        with open(temp_text_path, 'r', encoding='utf-8') as f:
                            text_content = f.read()
                    except Exception as e:
                        print(f"[DEBUG] Error reading text file: {str(e)}")

                    # Create a detailed error message with suggestions
                    error_message = "TEST RESULT - ERROR\n\n"
                    error_message += "No data was extracted. Here are some suggestions to fix your template:\n\n"

                    # Add specific suggestions based on the warnings
                    error_message += suggestions + "\n\n"

                    # Add general suggestions
                    error_message += "General suggestions:\n"
                    error_message += "1. Make sure your keywords match text in the invoice\n"
                    error_message += "2. Check that your regex patterns have capturing groups (parentheses)\n"
                    error_message += "3. Verify that the field names match those in the invoice\n"
                    error_message += "4. For date fields, ensure the formats list includes the format used in the invoice\n\n"

                    # Add a sample of the text content for reference
                    if text_content:
                        error_message += "Invoice text sample (first 500 chars):\n"
                        error_message += text_content[:500] + "...\n\n"

                    # Add technical details for debugging
                    error_message += "Technical details (for debugging):\n"
                    error_message += f"API Error: {str(api_error)}\n"
                    error_message += f"Command-line Error: {result.stderr}\n"

                    self.template_preview.setText(error_message)

                    # Show error message with suggestions
                    QMessageBox.warning(
                        self,
                        "Test Error",
                        "No data was extracted. See the template preview for suggestions on how to fix your template."
                    )

                except Exception as cmd_error:
                    # Both approaches failed
                    print(f"[DEBUG] Command-line fallback also failed: {str(cmd_error)}")

                    # Analyze the template and provide suggestions
                    suggestions = self.analyze_invoice2data_warnings("", template)

                    # Create a detailed error message with suggestions
                    error_message = "TEST RESULT - ERROR\n\n"
                    error_message += "No data was extracted. Here are some suggestions to fix your template:\n\n"

                    # Add specific suggestions based on template analysis
                    error_message += suggestions + "\n\n"

                    # Add general suggestions
                    error_message += "General suggestions:\n"
                    error_message += "1. Make sure your keywords match text in the invoice\n"
                    error_message += "2. Check that your regex patterns have capturing groups (parentheses)\n"
                    error_message += "3. Verify that the field names match those in the invoice\n"
                    error_message += "4. For date fields, ensure the formats list includes the format used in the invoice\n\n"

                    # Add technical details for debugging
                    error_message += "Technical details (for debugging):\n"
                    error_message += f"API Error: {str(api_error)}\n"
                    error_message += f"Command-line Error: {str(cmd_error)}\n"
                    error_message += f"API Traceback:\n{traceback_text}"

                    self.template_preview.setText(error_message)

                    # Show error message with suggestions
                    QMessageBox.warning(
                        self,
                        "Test Error",
                        "No data was extracted. See the template preview for suggestions on how to fix your template."
                    )

            # Clean up temporary files
            try:
                # Remove the text file
                if os.path.exists(temp_text_path):
                    os.remove(temp_text_path)
                    print(f"[DEBUG] Removed temporary text file: {temp_text_path}")

                # Remove the template file
                if os.path.exists(temp_template_path):
                    os.remove(temp_template_path)
                    print(f"[DEBUG] Removed temporary template file: {temp_template_path}")

                # Remove the temporary directory if it exists
                if 'temp_dir' in locals() and os.path.exists(temp_dir):
                    import shutil
                    shutil.rmtree(temp_dir)
                    print(f"[DEBUG] Removed temporary directory: {temp_dir}")
            except Exception as e:
                print(f"[DEBUG] Error cleaning up temporary files: {str(e)}")

        except Exception as e:
            print(f"[DEBUG] Error testing template: {str(e)}")
            import traceback
            traceback_text = traceback.format_exc()
            print(traceback_text)

            # Show detailed error in template preview
            error_message = f"TEST RESULT - ERROR\n\n"
            error_message += f"Error Type: {type(e).__name__}\n"
            error_message += f"Error Message: {str(e)}\n\n"
            error_message += "Traceback:\n"
            error_message += traceback_text

            self.template_preview.setText(error_message)

            # Show error message with the actual error
            QMessageBox.critical(
                self,
                "Test Error",
                f"An error occurred while testing the template:\n\n{type(e).__name__}: {str(e)}"
            )

        # Store the template for future use
        if hasattr(self, '_last_valid_template'):
            self._last_valid_template = template
            print("[DEBUG] Stored template as last valid template for future use")

    def save_invoice2data_template(self):
        """Save the invoice2data template to a file"""
        try:
            # Check if the template has all required fields
            required_fields = ['invoice_number', 'date', 'amount']
            existing_fields = []
            for row in range(self.fields_table.rowCount()):
                field_name = self.fields_table.item(row, 0).text().strip()
                existing_fields.append(field_name)

            missing_fields = [field for field in required_fields if field not in existing_fields]
            if missing_fields:
                # Show warning about missing fields
                warning_msg = f"The template is missing these required fields: {', '.join(missing_fields)}\n\n"
                warning_msg += "These fields are required by invoice2data. Do you want to continue anyway?"

                reply = QMessageBox.warning(
                    self,
                    "Missing Required Fields",
                    warning_msg,
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )

                if reply == QMessageBox.No:
                    return

            # Build the template
            template = self.build_invoice2data_template()

            # Get template name from issuer if not empty, otherwise prompt user
            template_name = template["issuer"].strip() if template["issuer"].strip() else None
            if not template_name:
                template_name, ok = QInputDialog.getText(self, "Save Template", "Enter template name:")
                if not ok or not template_name.strip():
                    return

            # Sanitize template name for filename
            template_name = re.sub(r'[^\w\-_.]', '_', template_name)

            # Ask user for save location and format
            formats = ["YAML (.yml)", "JSON (.json)"]
            selected_format, ok = QInputDialog.getItem(
                self, "Save Format", "Select format:", formats, 0, False
            )
            if not ok:
                return

            # Get file extension based on selected format
            file_ext = ".json" if "JSON" in selected_format else ".yml"

            # Ask user for save location
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Template",
                os.path.join(os.path.expanduser("~"), f"{template_name}{file_ext}"),
                f"Template Files (*{file_ext});;All Files (*.*)"
            )

            if not save_path:
                return

            # Ensure the file has the correct extension
            if not save_path.endswith(file_ext):
                save_path += file_ext

            # Add required fields for invoice2data
            if "name" not in template:
                template["name"] = template_name
            if "template_name" not in template:
                template["template_name"] = template_name

            # Save the template to the selected file - ensure regex patterns are preserved
            if file_ext == ".json":
                with open(save_path, "w", encoding="utf-8") as f:
                    # Use ensure_ascii=False to preserve special characters
                    json.dump(template, f, indent=2, ensure_ascii=False)
            else:
                # Convert to YAML and save
                import yaml
                with open(save_path, "w", encoding="utf-8") as f:
                    yaml.dump(template, f, default_flow_style=False, allow_unicode=True)

            # Show success message
            QMessageBox.information(
                self,
                "Template Saved",
                f"Template saved to {save_path}"
            )

            # Emit signal that template was created
            self.invoice2data_template_created.emit(template)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save template: {str(e)}")

    def analyze_invoice2data_warnings(self, warnings_text, template):
        """Analyze invoice2data warnings and suggest fixes

        Args:
            warnings_text (str): The warnings captured from invoice2data
            template (dict): The template that was used

        Returns:
            str: Suggestions for fixing the warnings
        """
        # Use the unified invoice processing utilities
        return invoice_processing_utils.analyze_invoice2data_warnings(warnings_text, template)

    def process_with_invoice2data(self, pdf_path=None, template_data=None, extracted_data=None, temp_dir=None):
        """Process a PDF file with invoice2data using the template and extracted data

        Args:
            pdf_path (str, optional): Path to the PDF file. If None, uses self.pdf_path
            template_data (dict, optional): Template data from the database. If None, uses current template
            extracted_data (dict, optional): Extracted data from the PDF. If None, uses current extraction data
            temp_dir (str, optional): Path to temporary directory

        Returns:
            dict: The extraction result from invoice2data, or None if extraction failed
        """
        # Use the unified invoice_processing_utils for invoice2data processing
        try:
            # Use default values if not provided
            if pdf_path is None:
                pdf_path = self.pdf_path

            if template_data is None:
                # Use the current template if available
                if hasattr(self, '_last_valid_template') and self._last_valid_template:
                    template_data = {'json_template': self._last_valid_template}
                else:
                    # Build a new template from the form
                    template_data = {'json_template': self.build_invoice2data_template()}

            if extracted_data is None:
                # Use the current extraction data
                extracted_data = self._get_current_json_data()

            # Use the unified invoice processing utilities
            print(f"[DEBUG] Using invoice_processing_utils.process_with_invoice2data")
            return invoice_processing_utils.process_with_invoice2data(
                pdf_path=pdf_path,
                template_data=template_data,
                extracted_data=extracted_data,
                temp_dir=temp_dir
            )
        except Exception as e:
            print(f"[DEBUG] Error in process_with_invoice2data: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def convert_extraction_to_text(self, extraction_data):
        """Convert the extraction data to a text format that can be used by invoice2data

        Args:
            extraction_data (dict): The extracted data from pypdf_table_extraction

        Returns:
            str: Text representation of the extracted data with pipe-separated values
        """
        # Use the unified invoice processing utilities with pdf_path parameter
        return invoice_processing_utils.convert_extraction_to_text(extraction_data, pdf_path=self.pdf_path if hasattr(self, 'pdf_path') else None)

    # The clear_all method has been removed as its functionality is now merged into reset_screen

    def reset_screen(self):
        """Reset the screen to its initial state, clear cache, and clean temp directory"""
        print(f"[DEBUG] Resetting screen")

        # Reset regions and column lines
        self.regions = {'header': [], 'items': [], 'summary': []}
        self.column_lines = {RegionType.HEADER: [], RegionType.ITEMS: [], RegionType.SUMMARY: []}

        # Reset drawing state
        self.current_region_type = None
        self.drawing_column = False

        # Reset cursor to default
        if self.pdf_label:
            self.pdf_label.setCursor(Qt.ArrowCursor)

        # Reset multi-page mode
        self.multi_page_mode = False
        self.current_page_index = 0

        # Reset splitter sizes with smaller JSON designer section
        total_width = self.main_splitter.width()
        # Use a 4:4:2 ratio to make the JSON designer section smaller
        self.main_splitter.setSizes([int(total_width * 0.4), int(total_width * 0.4), int(total_width * 0.2)])

        # Reset the user adjusted splitter flag
        self._user_adjusted_splitter = False
        self.page_regions = {}
        self.page_column_lines = {}

        # Hide navigation buttons
        self.prev_page_btn.hide()
        self.next_page_btn.hide()
        self.apply_to_remaining_btn.hide()

        # Uncheck all buttons
        self.header_btn.setChecked(False)
        self.items_btn.setChecked(False)
        self.summary_btn.setChecked(False)
        self.column_btn.setChecked(False)

        # Clear cached extraction data
        self._cached_extraction_data = {
            'header': [],
            'items': [],
            'summary': []
        }
        self._last_extraction_state = None

        # Reset PDF display state with proper resource cleanup
        if hasattr(self, 'pdf_document') and self.pdf_document:
            try:
                self.pdf_document.close()
                print("[DEBUG] Closed PDF document during reset")
            except Exception as e:
                print(f"[WARNING] Error closing PDF document during reset: {e}")
            finally:
                self.pdf_document = None

        self.pdf_path = None

        # Hide zoom controls
        if hasattr(self, 'zoom_controls'):
            self.zoom_controls.hide()

        # Clear the PDF label pixmaps and show upload area
        if hasattr(self, 'pdf_label'):
            # Clear both original and scaled pixmaps to free memory
            if hasattr(self.pdf_label, 'original_pixmap'):
                self.pdf_label.original_pixmap = None
            if hasattr(self.pdf_label, 'scaled_pixmap'):
                self.pdf_label.scaled_pixmap = None
            self.pdf_label.clear()
            self.pdf_label.hide()
            print("[DEBUG] Cleared PDF label pixmaps during reset")

        if hasattr(self, 'upload_area'):
            self.upload_area.show()

        # Hide PDF controls when clearing everything
        self.pdf_controls_container.hide()

        # Hide scrollbars when no PDF is loaded
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        print(f"[DEBUG] Hiding scrollbars when no PDF is loaded")

        # Update the JSON tree
        self.update_json_tree(self._cached_extraction_data)

        # Reset the JSON designer screen
        self.reset_json_designer()

        # Clear memory cache
        import gc
        gc.collect()

        # Clear extraction cache
        if hasattr(self, 'pdf_path') and self.pdf_path:
            clear_extraction_cache_for_pdf(self.pdf_path)
            print(f"[DEBUG] Cleared extraction cache for PDF: {self.pdf_path}")

        # Clean up templates directory
        if PATH_HELPER_AVAILABLE:
            templates_dir = path_helper.resolve_path("templates")
        else:
            templates_dir = os.path.abspath("templates")
        if os.path.exists(templates_dir):
            for file in os.listdir(templates_dir):
                if file.startswith("test") and (file.endswith(".json") or file.endswith(".yml") or file.endswith(".txt")):
                    try:
                        os.remove(os.path.join(templates_dir, file))
                        print(f"[DEBUG] Removed temporary file: {os.path.join(templates_dir, file)}")
                    except Exception as e:
                        print(f"[ERROR] Failed to remove file {file}: {str(e)}")

        # Clear all caches to free memory
        self._clear_all_caches()

        # Force garbage collection to free memory
        import gc
        collected = gc.collect()
        print(f"[DEBUG] Reset screen completed successfully, freed {collected} objects")

        # Update the display
        self.pdf_label.update()

        # Show success message
        QMessageBox.information(
            self,
            "Reset Complete",
            "Screen has been reset, cache cleared, and temp files removed."
        )

    def handle_go_back(self):
        """Handle going back to the previous screen"""
        print(f"[DEBUG] handle_go_back method called")

        # Clean up resources before going back
        if hasattr(self, 'pdf_document') and self.pdf_document:
            try:
                self.pdf_document.close()
                self.pdf_document = None
                print(f"[DEBUG] Closed PDF document")
            except Exception as e:
                print(f"[DEBUG] Error closing PDF document: {str(e)}")

        # Reset UI state
        if hasattr(self, 'pdf_label'):
            self.pdf_label.clear()
            self.pdf_label.hide()
            print(f"[DEBUG] Cleared and hid PDF label")

        if hasattr(self, 'upload_area'):
            self.upload_area.show()
            print(f"[DEBUG] Showed upload area")

        # Emit the go_back signal directly
        self.go_back.emit()

    # Zoom methods
    def zoom_in(self):
        """Zoom in on the PDF"""
        try:
            # Make sure the PDF label exists
            if not hasattr(self, 'pdf_label') or self.pdf_label is None:
                print("[DEBUG] Cannot zoom in: PDF label not found")
                return

            # Initialize zoom level if it doesn't exist
            if not hasattr(self.pdf_label, 'zoom_level'):
                self.pdf_label.zoom_level = 1.0

            # Get current zoom level before changing
            old_zoom = self.pdf_label.zoom_level

            # Use smaller zoom factor for smoother zooming
            self.pdf_label.zoom_level *= 1.05
            self.pdf_label.zoom_level = min(2.5, self.pdf_label.zoom_level)  # Limit max zoom

            # Only proceed if zoom actually changed
            if abs(old_zoom - self.pdf_label.zoom_level) > 0.01:
                # Update the zoom label
                self.update_zoom_label(self.pdf_label.zoom_level)
                print(f"[DEBUG] Zoomed in to: {self.pdf_label.zoom_level:.2f}x")
        except Exception as e:
            print(f"[DEBUG] Error in zoom_in: {str(e)}")
            import traceback
            traceback.print_exc()

    def ensure_full_scroll_range(self):
        """Ensure the scroll area can scroll to the end of the document"""
        if not hasattr(self, 'pdf_label') or not hasattr(self, 'scroll_area'):
            return

        # Get the current size of the PDF label
        label_size = self.pdf_label.size()

        # Get the viewport size and use it for calculating padding
        viewport_size = self.scroll_area.viewport().size()
        print(f"[DEBUG] Viewport size: {viewport_size.width()}x{viewport_size.height()}")

        # Get current zoom level
        zoom_level = getattr(self.pdf_label, 'zoom_level', 1.0)

        # Calculate extra padding based on zoom level and viewport size
        extra_padding = int(max(500, viewport_size.height() * 0.5) * zoom_level)  # More padding at higher zoom levels

        # Always add extra space to ensure scrolling works, especially when zoomed in
        new_height = label_size.height() + extra_padding
        self.pdf_label.setFixedHeight(new_height)
        print(f"[DEBUG] Added extra space to PDF label for scrolling: {label_size.height()} -> {new_height} (zoom: {zoom_level:.2f}x)")

        # Force the scroll area to update
        self.scroll_area.updateGeometry()
        QApplication.processEvents()

    def zoom_out(self):
        """Zoom out on the PDF"""
        try:
            # Make sure the PDF label exists
            if not hasattr(self, 'pdf_label') or self.pdf_label is None:
                print("[DEBUG] Cannot zoom out: PDF label not found")
                return

            # Initialize zoom level if it doesn't exist
            if not hasattr(self.pdf_label, 'zoom_level'):
                self.pdf_label.zoom_level = 1.0

            # Get current zoom level before changing
            old_zoom = self.pdf_label.zoom_level

            # Use smaller zoom factor for smoother zooming
            self.pdf_label.zoom_level *= 0.95
            self.pdf_label.zoom_level = max(0.5, self.pdf_label.zoom_level)  # Limit min zoom

            # Only proceed if zoom actually changed
            if abs(old_zoom - self.pdf_label.zoom_level) > 0.01:
                # Update the zoom label
                self.update_zoom_label(self.pdf_label.zoom_level)
                print(f"[DEBUG] Zoomed out to: {self.pdf_label.zoom_level:.2f}x")
        except Exception as e:
            print(f"[DEBUG] Error in zoom_out: {str(e)}")
            import traceback
            traceback.print_exc()

    def reset_zoom(self):
        """Reset zoom to default level"""
        try:
            # Make sure the PDF label exists
            if not hasattr(self, 'pdf_label') or self.pdf_label is None:
                print("[DEBUG] Cannot reset zoom: PDF label not found")
                return

            # Reset zoom level to 1.0 (100%)
            self.pdf_label.zoom_level = 1.0

            # Update the zoom label
            self.update_zoom_label(1.0)
            print(f"[DEBUG] Reset zoom to 100%")
        except Exception as e:
            print(f"[DEBUG] Error in reset_zoom: {str(e)}")
            import traceback
            traceback.print_exc()

    def create_floating_zoom_controls(self):
        """Create floating zoom controls that appear in the middle of the PDF viewer"""
        # Create a widget to hold the zoom controls
        self.zoom_controls = QWidget(self.scroll_area.viewport())
        self.zoom_controls.setObjectName("zoomControls")

        # Set up the layout
        zoom_layout = QHBoxLayout(self.zoom_controls)
        zoom_layout.setContentsMargins(5, 5, 5, 5)
        zoom_layout.setSpacing(2)

        # Zoom out button
        self.zoom_out_btn = QPushButton("-")
        self.zoom_out_btn.setFixedSize(30, 30)
        self.zoom_out_btn.clicked.connect(self.zoom_out)
        self.zoom_out_btn.setCursor(Qt.PointingHandCursor)
        zoom_layout.addWidget(self.zoom_out_btn)

        # Zoom level indicator
        self.zoom_label = QLabel("100%")
        self.zoom_label.setAlignment(Qt.AlignCenter)
        self.zoom_label.setFixedWidth(60)
        zoom_layout.addWidget(self.zoom_label)

        # Zoom in button
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setFixedSize(30, 30)
        self.zoom_in_btn.clicked.connect(self.zoom_in)
        self.zoom_in_btn.setCursor(Qt.PointingHandCursor)
        zoom_layout.addWidget(self.zoom_in_btn)

        # Reset zoom button
        self.reset_zoom_btn = QPushButton("Reset")
        self.reset_zoom_btn.setFixedHeight(30)
        self.reset_zoom_btn.clicked.connect(self.reset_zoom)
        self.reset_zoom_btn.setCursor(Qt.PointingHandCursor)
        zoom_layout.addWidget(self.reset_zoom_btn)

        # Style the zoom controls
        self.zoom_controls.setStyleSheet("""
            #zoomControls {
                background-color: rgba(0, 0, 0, 220);
                border: 1px solid #333333;
                border-radius: 15px;
            }
            QPushButton {
                background-color: rgba(40, 40, 40, 220);
                border: 1px solid #444444;
                border-radius: 15px;
                font-weight: bold;
                font-size: 16px;
                color: white;
            }
            QPushButton:hover {
                background-color: rgba(60, 60, 60, 240);
                border-color: #555555;
            }
            QPushButton#reset_zoom_btn {
                border-radius: 10px;
                font-size: 12px;
            }
            QLabel {
                background-color: rgba(40, 40, 40, 220);
                border: 1px solid #444444;
                border-radius: 10px;
                padding: 2px;
                color: white;
                font-weight: bold;
            }
        """)

        # Set initial size and position
        self.zoom_controls.setFixedSize(180, 40)
        self.zoom_controls.hide()  # Initially hidden

        # Install event filter on scroll area viewport to reposition controls when scrolling
        self.scroll_area.viewport().installEventFilter(self)

        # Position the controls at the bottom center of the scroll area
        self.position_zoom_controls()

        # Show the zoom controls when the PDF is loaded
        self.zoom_controls.show()

        # Make sure the zoom controls stay on top
        self.zoom_controls.raise_()

    def position_zoom_controls(self):
        """Position the zoom controls at the bottom center of the scroll area"""
        if not hasattr(self, 'zoom_controls') or not hasattr(self, 'scroll_area'):
            return

        try:
            # Get the scroll area viewport size
            viewport_size = self.scroll_area.viewport().size()

            # Calculate position (centered horizontally, near the bottom)
            x = max(0, (viewport_size.width() - self.zoom_controls.width()) // 2)
            y = max(0, viewport_size.height() - self.zoom_controls.height() - 20)  # 20px from bottom

            # Make sure the zoom controls are a child of the viewport
            if self.zoom_controls.parent() != self.scroll_area.viewport():
                self.zoom_controls.setParent(self.scroll_area.viewport())

            # Move the controls to the calculated position
            self.zoom_controls.move(x, y)

            # Make sure the controls are visible
            self.zoom_controls.show()

            # Ensure the controls stay on top
            self.zoom_controls.raise_()

            # print(f"[DEBUG] Positioned zoom controls at ({x}, {y}) in viewport of size {viewport_size.width()}x{viewport_size.height()}")
        except Exception as e:
            print(f"[DEBUG] Error positioning zoom controls: {str(e)}")
            import traceback
            traceback.print_exc()

    def update_zoom_label(self, zoom_level):
        """Update the zoom level label"""
        try:
            # Calculate percentage from zoom level
            percentage = int(zoom_level * 100)

            # Update the zoom label text if it exists
            if hasattr(self, 'zoom_label') and self.zoom_label is not None:
                self.zoom_label.setText(f"{percentage}%")
                print(f"[DEBUG] Updated zoom label to {percentage}%")

                # Make sure the zoom controls are visible
                if hasattr(self, 'zoom_controls') and self.zoom_controls is not None:
                    # Show the zoom controls if they're hidden
                    if not self.zoom_controls.isVisible():
                        self.zoom_controls.show()

                    # Reposition zoom controls when zoom changes
                    self.position_zoom_controls()

                    # Make sure the zoom controls stay on top
                    self.zoom_controls.raise_()
            else:
                print(f"[DEBUG] zoom_label not found")

            # Update the PDF display to reflect the new zoom level
            if hasattr(self, 'pdf_label') and self.pdf_label is not None:
                # Ensure the PDF label's zoom level matches
                self.pdf_label.zoom_level = zoom_level

                # Adjust the pixmap to reflect the new zoom level
                self.pdf_label.adjustPixmap()

                # Ensure proper scrolling after zoom change
                self.ensure_full_scroll_range()

        except Exception as e:
            print(f"[DEBUG] Error updating zoom label: {str(e)}")
            import traceback
            traceback.print_exc()

    # Event filter for handling scroll events
    def eventFilter(self, obj, event):
        """Event filter to handle scroll events"""
        try:
            if obj == self.scroll_area.viewport():
                # Handle different event types
                if event.type() == QEvent.Resize:
                    # Viewport was resized, reposition zoom controls
                    if hasattr(self, 'zoom_controls') and self.zoom_controls.isVisible():
                        self.position_zoom_controls()

                elif event.type() == QEvent.Wheel:
                    # Mouse wheel event in the viewport - could be zooming
                    # Let the event propagate to the PDF label for zooming
                    pass

                elif event.type() in [QEvent.Scroll, QEvent.Paint]:
                    # Scrolling or painting event, reposition zoom controls
                    if hasattr(self, 'zoom_controls') and self.zoom_controls.isVisible():
                        self.position_zoom_controls()

            # Special handling for wheel events on the PDF label
            elif hasattr(self, 'pdf_label') and obj == self.pdf_label and event.type() == QEvent.Wheel:
                # Let the PDF label handle the wheel event for zooming
                # The zoom_changed signal will be emitted and update_zoom_label will be called
                pass

        except Exception as e:
            print(f"[DEBUG] Error in eventFilter: {str(e)}")
            import traceback
            traceback.print_exc()

        # Always return False to allow the event to be processed further
        return False

    # Window resize event
    def resizeEvent(self, event):
        """Handle window resize event"""
        super().resizeEvent(event)

        # Reposition zoom controls when window is resized
        if hasattr(self, 'zoom_controls') and self.zoom_controls.isVisible():
            self.position_zoom_controls()

        # Maintain the 4:4:2 ratio for the splitter sections
        self.maintain_splitter_ratio()

    def maintain_splitter_ratio(self):
        """Maintain the 4:4:2 ratio for the splitter sections"""
        # Only maintain ratio if not manually adjusted by user
        if not hasattr(self, '_user_adjusted_splitter') or not self._user_adjusted_splitter:
            # Get the total width of the splitter
            total_width = self.main_splitter.width()

            # Check if any section is hidden (width < 10)
            sizes = self.main_splitter.sizes()
            if all(size >= 10 for size in sizes):
                # Set sizes to maintain 4:4:2 ratio
                self.main_splitter.setSizes([int(total_width * 0.4), int(total_width * 0.4), int(total_width * 0.2)])
                print(f"[DEBUG] Maintaining 4:4:2 ratio for splitter sections: {self.main_splitter.sizes()}")

    def on_splitter_moved(self, pos, index):
        """Handle splitter movement to detect when PDF section visibility changes"""
        # Mark that the user has manually adjusted the splitter
        self._user_adjusted_splitter = True
        print(f"[DEBUG] Splitter moved to position {pos} at index {index}")

        # Check if the PDF section (index 0) is now visible after being hidden
        pdf_section_width = self.main_splitter.sizes()[0]

        # Check if PDF section was hidden and is now visible
        if self.pdf_section_was_hidden and pdf_section_width > 10:
            print(f"[DEBUG] PDF section was hidden and is now visible. Refreshing PDF display.")
            self.pdf_section_was_hidden = False

            # Refresh the PDF display
            if hasattr(self, 'pdf_document') and self.pdf_document and hasattr(self, 'pdf_label'):
                # Force a redraw of the PDF
                self.display_current_page()

                # Make sure the PDF label is visible
                self.pdf_label.show()

                # Ensure zoom controls are properly positioned
                if hasattr(self, 'zoom_controls') and self.zoom_controls.isVisible():
                    self.position_zoom_controls()

        # Check if PDF section is now hidden
        elif pdf_section_width < 10:
            self.pdf_section_was_hidden = True
            print(f"[DEBUG] PDF section is now hidden.")

    # Drag and drop support
    def dragEnterEvent(self, event):
        """Handle drag enter event"""
        if event.mimeData().hasUrls() and any(url.toLocalFile().endswith('.pdf') for url in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(self, event):
        """Handle drop event"""
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith('.pdf'):
                self.load_pdf(file_path)
                break

    def get_template_page_index(self, pdf_page_index):
        """Determine which template page to use for a given PDF page

        Args:
            pdf_page_index (int): The 0-based index of the PDF page

        Returns:
            int: The 0-based index of the template page to use
        """
        try:
            # Use unified mapping logic if available
            if INVOICE_PROCESSING_UTILS_AVAILABLE:
                # Create template data for mapping
                template_data = {
                    'template_type': 'multi' if getattr(self, 'multi_page_mode', False) else 'single',
                    'page_count': len(getattr(self, 'page_configs', [])) if hasattr(self, 'page_configs') else 1,
                    'config': {
                        'use_middle_page': getattr(self, 'use_middle_page', False),
                        'fixed_page_count': getattr(self, 'fixed_page_count', False)
                    }
                }

                pdf_total_pages = len(self.pdf_document) if self.pdf_document else 1

                template_page_index = invoice_processing_utils.get_template_page_for_pdf_page(
                    pdf_page_index=pdf_page_index,
                    pdf_total_pages=pdf_total_pages,
                    template_data=template_data
                )

                print(f"[DEBUG] Unified mapping: PDF page {pdf_page_index + 1} → Template page {template_page_index + 1}")
                return template_page_index

        except Exception as e:
            print(f"[WARNING] Could not use unified mapping logic: {e}")

        # Fallback logic
        if not getattr(self, 'multi_page_mode', False):
            # Single-page mode: always use template page 0
            return 0

        # Multi-page mode fallback logic
        if hasattr(self, 'page_configs') and self.page_configs:
            template_page_count = len(self.page_configs)

            # Simple mapping: use the same index, but don't exceed template page count
            template_page_index = min(pdf_page_index, template_page_count - 1)
            print(f"[DEBUG] Fallback mapping: PDF page {pdf_page_index + 1} → Template page {template_page_index + 1}")
            return template_page_index
        else:
            # No page configs available, use page 0
            print(f"[DEBUG] No page configs available, using template page 1 for PDF page {pdf_page_index + 1}")
            return 0

    def cleanup_processor_caches(self):
        """Clean up all caches and temporary data for this processor"""
        try:
            print("[DEBUG] Starting split_screen_invoice_processor cache cleanup...")

            # Clear extraction cache if a PDF was loaded
            if hasattr(self, 'pdf_path') and self.pdf_path:
                clear_extraction_cache_for_pdf(self.pdf_path)
                print(f"[DEBUG] Cleared extraction cache for PDF: {self.pdf_path}")

            # Clear all extraction caches
            clear_extraction_cache()
            print("[DEBUG] Cleared all extraction caches")

            # Clear instance-specific cached data
            if hasattr(self, '_cached_extraction_data'):
                self._cached_extraction_data = None
                print("[DEBUG] Cleared cached extraction data")

            if hasattr(self, '_all_pages_data'):
                self._all_pages_data = None
                print("[DEBUG] Cleared all pages data")

            # Clear region data
            if hasattr(self, 'regions'):
                self.regions = {'header': [], 'items': [], 'summary': []}
            if hasattr(self, 'page_regions'):
                self.page_regions = {}
            if hasattr(self, 'page_column_lines'):
                self.page_column_lines = {}
            print("[DEBUG] Cleared region data")

            # Clean up templates directory
            if PATH_HELPER_AVAILABLE:
                templates_dir = path_helper.resolve_path("templates")
            else:
                templates_dir = os.path.abspath("templates")
            if os.path.exists(templates_dir):
                temp_files_removed = 0
                for file in os.listdir(templates_dir):
                    if file.startswith("test") and (file.endswith(".json") or file.endswith(".yml") or file.endswith(".txt")):
                        try:
                            os.remove(os.path.join(templates_dir, file))
                            temp_files_removed += 1
                            print(f"[DEBUG] Removed temporary file: {os.path.join(templates_dir, file)}")
                        except Exception as e:
                            print(f"[ERROR] Failed to remove file {file}: {str(e)}")

                if temp_files_removed > 0:
                    print(f"[DEBUG] Removed {temp_files_removed} temporary template files")

            # Force garbage collection
            import gc
            collected = gc.collect()
            print(f"[DEBUG] Garbage collection: {collected} objects collected")

            print("[DEBUG] Split_screen_invoice_processor cache cleanup completed")

        except Exception as e:
            print(f"[ERROR] Error in cleanup_processor_caches: {str(e)}")
            import traceback
            traceback.print_exc()

    def closeEvent(self, event):
        """Clean up resources when the window is closed"""
        try:
            print("[DEBUG] SplitScreenInvoiceProcessor closeEvent triggered")

            # Perform comprehensive cache cleanup
            self.cleanup_processor_caches()

            # Register with cache manager if available
            if CACHE_MANAGER_AVAILABLE:
                try:
                    cache_manager = get_cache_manager()
                    # Run any additional cleanup through cache manager
                    cache_manager.clear_extraction_caches()
                    print("[DEBUG] Cache manager cleanup completed")
                except Exception as e:
                    print(f"[WARNING] Cache manager cleanup failed: {e}")

            # Call the parent class closeEvent
            super().closeEvent(event)

        except Exception as e:
            print(f"[ERROR] Error in closeEvent: {str(e)}")
            import traceback
            traceback.print_exc()
            # Make sure the event is accepted even if there's an error
            event.accept()
