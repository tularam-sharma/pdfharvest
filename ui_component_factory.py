"""
UI Component Factory

This module provides factory functions to create standardized UI components
and eliminate duplicate UI initialization patterns across the application.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QProgressBar,
    QFrame, QGroupBox, QLineEdit, QTextEdit, QComboBox,
    QCheckBox, QSpinBox, QTabWidget, QSplitter, QStackedWidget
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QColor
from typing import List, Dict, Optional, Any, Callable


class UIComponentFactory:
    """Factory for creating standardized UI components"""
    
    @staticmethod
    def create_standard_button(
        text: str,
        callback: Callable = None,
        style_type: str = "primary",
        enabled: bool = True
    ) -> QPushButton:
        """Create a standardized button with consistent styling"""
        button = QPushButton(text)
        button.setEnabled(enabled)
        
        if callback:
            button.clicked.connect(callback)
        
        # Apply standard styling based on type
        styles = {
            "primary": """
                QPushButton {
                    background-color: #007bff;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 4px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                }
                QPushButton:disabled {
                    background-color: #6c757d;
                }
            """,
            "secondary": """
                QPushButton {
                    background-color: #6c757d;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 4px;
                    font-weight: normal;
                }
                QPushButton:hover {
                    background-color: #545b62;
                }
            """,
            "danger": """
                QPushButton {
                    background-color: #dc3545;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 4px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #c82333;
                }
            """,
            "success": """
                QPushButton {
                    background-color: #28a745;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 4px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #218838;
                }
            """
        }
        
        button.setStyleSheet(styles.get(style_type, styles["primary"]))
        return button
    
    @staticmethod
    def create_standard_table(
        headers: List[str],
        data: List[List[str]] = None,
        sortable: bool = True,
        selectable: bool = True
    ) -> QTableWidget:
        """Create a standardized table widget"""
        table = QTableWidget()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        
        # Configure table properties
        if sortable:
            table.setSortingEnabled(True)
        
        if selectable:
            table.setSelectionBehavior(QTableWidget.SelectRows)
            table.setSelectionMode(QTableWidget.SingleSelection)
        
        # Set header properties
        header = table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(QHeaderView.Interactive)
        
        # Add data if provided
        if data:
            table.setRowCount(len(data))
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_data in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_data))
                    table.setItem(row_idx, col_idx, item)
        
        # Apply standard styling
        table.setStyleSheet("""
            QTableWidget {
                gridline-color: #d1d1d1;
                background-color: white;
                alternate-background-color: #f8f9fa;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #dee2e6;
            }
            QTableWidget::item:selected {
                background-color: #007bff;
                color: white;
            }
            QHeaderView::section {
                background-color: #f8f9fa;
                padding: 8px;
                border: 1px solid #dee2e6;
                font-weight: bold;
            }
        """)
        
        return table
    
    @staticmethod
    def create_standard_progress_bar(
        minimum: int = 0,
        maximum: int = 100,
        value: int = 0,
        show_text: bool = True
    ) -> QProgressBar:
        """Create a standardized progress bar"""
        progress = QProgressBar()
        progress.setMinimum(minimum)
        progress.setMaximum(maximum)
        progress.setValue(value)
        progress.setTextVisible(show_text)
        
        # Apply standard styling
        progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #dee2e6;
                border-radius: 5px;
                text-align: center;
                background-color: #f8f9fa;
            }
            QProgressBar::chunk {
                background-color: #007bff;
                border-radius: 3px;
            }
        """)
        
        return progress
    
    @staticmethod
    def create_standard_group_box(
        title: str,
        layout_type: str = "vertical"
    ) -> QGroupBox:
        """Create a standardized group box with layout"""
        group_box = QGroupBox(title)
        
        if layout_type == "vertical":
            layout = QVBoxLayout()
        elif layout_type == "horizontal":
            layout = QHBoxLayout()
        else:
            layout = QVBoxLayout()  # Default to vertical
        
        group_box.setLayout(layout)
        
        # Apply standard styling
        group_box.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                background-color: white;
            }
        """)
        
        return group_box
    
    @staticmethod
    def create_standard_input_field(
        placeholder: str = "",
        input_type: str = "text",
        max_length: int = None
    ) -> QWidget:
        """Create a standardized input field"""
        if input_type == "text":
            widget = QLineEdit()
            widget.setPlaceholderText(placeholder)
            if max_length:
                widget.setMaxLength(max_length)
        elif input_type == "multiline":
            widget = QTextEdit()
            widget.setPlaceholderText(placeholder)
        elif input_type == "number":
            widget = QSpinBox()
            widget.setMinimum(0)
            widget.setMaximum(999999)
        elif input_type == "combo":
            widget = QComboBox()
        else:
            widget = QLineEdit()  # Default to text
        
        # Apply standard styling
        widget.setStyleSheet("""
            QLineEdit, QTextEdit, QSpinBox, QComboBox {
                padding: 8px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                background-color: white;
            }
            QLineEdit:focus, QTextEdit:focus, QSpinBox:focus, QComboBox:focus {
                border-color: #007bff;
                outline: none;
            }
        """)
        
        return widget
    
    @staticmethod
    def create_button_layout(
        buttons: List[Dict[str, Any]],
        layout_type: str = "horizontal"
    ) -> QWidget:
        """Create a layout with multiple standardized buttons"""
        container = QWidget()
        
        if layout_type == "horizontal":
            layout = QHBoxLayout()
        else:
            layout = QVBoxLayout()
        
        for button_config in buttons:
            button = UIComponentFactory.create_standard_button(
                text=button_config.get("text", "Button"),
                callback=button_config.get("callback"),
                style_type=button_config.get("style", "primary"),
                enabled=button_config.get("enabled", True)
            )
            layout.addWidget(button)
        
        # Add stretch to push buttons to one side
        if layout_type == "horizontal":
            layout.addStretch()
        
        container.setLayout(layout)
        return container
    
    @staticmethod
    def create_standard_splitter(
        orientation: str = "horizontal",
        sizes: List[int] = None
    ) -> QSplitter:
        """Create a standardized splitter"""
        if orientation == "horizontal":
            splitter = QSplitter(Qt.Horizontal)
        else:
            splitter = QSplitter(Qt.Vertical)
        
        if sizes:
            splitter.setSizes(sizes)
        
        # Apply standard styling
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #dee2e6;
                border: 1px solid #ced4da;
            }
            QSplitter::handle:horizontal {
                width: 3px;
            }
            QSplitter::handle:vertical {
                height: 3px;
            }
        """)
        
        return splitter


class LayoutFactory:
    """Factory for creating standardized layouts"""
    
    @staticmethod
    def create_form_layout(
        fields: List[Dict[str, Any]],
        parent: QWidget = None
    ) -> QWidget:
        """Create a standardized form layout"""
        container = QWidget(parent)
        layout = QVBoxLayout()
        
        for field_config in fields:
            field_layout = QHBoxLayout()
            
            # Create label
            label = QLabel(field_config.get("label", "Field:"))
            label.setMinimumWidth(120)
            field_layout.addWidget(label)
            
            # Create input field
            input_field = UIComponentFactory.create_standard_input_field(
                placeholder=field_config.get("placeholder", ""),
                input_type=field_config.get("type", "text"),
                max_length=field_config.get("max_length")
            )
            field_layout.addWidget(input_field)
            
            layout.addLayout(field_layout)
        
        container.setLayout(layout)
        return container


# Global factory instances
ui_factory = UIComponentFactory()
layout_factory = LayoutFactory()
