"""
Common Factories Module

This module provides factory classes to eliminate duplicate code patterns
across the PDF_EXTRACTOR application. It centralizes common creation logic
for templates, database operations, and UI components.
"""

import json
from typing import Dict, List, Optional, Any, Union
from datetime import datetime
from PySide6.QtWidgets import QMessageBox, QWidget
from database import InvoiceDatabase


class TemplateFactory:
    """Factory for creating standardized templates"""
    
    @staticmethod
    def create_default_invoice_template(name: str, currency: str = "INR") -> Dict[str, Any]:
        """Create a default invoice template with standard structure"""
        return {
            "issuer": name,
            "fields": {
                "invoice_number": "1",
                "date": "1",
                "amount": "1"
            },
            "keywords": [name],
            "options": {
                "currency": currency,
                "languages": ["en"],
                "decimal_separator": ".",
                "replace": []
            }
        }
    
    @staticmethod
    def create_default_regions() -> Dict[str, List]:
        """Create default empty regions structure"""
        return {
            "header": [],
            "items": [],
            "summary": []
        }
    
    @staticmethod
    def create_default_column_lines() -> Dict[str, List]:
        """Create default empty column lines structure"""
        return {
            "header": [],
            "items": [],
            "summary": []
        }
    
    @staticmethod
    def create_default_config() -> Dict[str, Any]:
        """Create default extraction configuration"""
        return {
            "multi_table_mode": False,
            "extraction_params": {
                "header": {"row_tol": 5},
                "items": {"row_tol": 5},
                "summary": {"row_tol": 5},
                "split_text": True,
                "strip_text": "\n",
                "flavor": "stream"
            }
        }
    
    @staticmethod
    def create_complete_template(
        name: str,
        description: str = "",
        template_type: str = "single",
        currency: str = "INR"
    ) -> Dict[str, Any]:
        """Create a complete template with all required components"""
        return {
            "name": name,
            "description": description or f"Template for {name}",
            "template_type": template_type,
            "regions": TemplateFactory.create_default_regions(),
            "column_lines": TemplateFactory.create_default_column_lines(),
            "config": TemplateFactory.create_default_config(),
            "json_template": TemplateFactory.create_default_invoice_template(name, currency)
        }


class DatabaseOperationFactory:
    """Factory for standardized database operations"""
    
    def __init__(self, db: InvoiceDatabase):
        self.db = db
    
    def save_template_safe(
        self,
        name: str,
        description: str = "",
        template_type: str = "single",
        currency: str = "INR",
        **kwargs
    ) -> Optional[int]:
        """Safely save a template with validation and error handling"""
        try:
            # Check if template already exists
            existing = self.db.get_template(template_name=name)
            if existing:
                return None  # Template exists
            
            # Create template using factory
            template_data = TemplateFactory.create_complete_template(
                name, description, template_type, currency
            )
            
            # Override with any provided kwargs
            template_data.update(kwargs)
            
            # Save to database with dual coordinate support
            template_id = self.db.save_template(
                name=template_data["name"],
                description=template_data["description"],
                config=template_data["config"],
                template_type=template_data["template_type"],
                json_template=template_data["json_template"],
                drawing_regions=template_data.get("drawing_regions"),
                drawing_column_lines=template_data.get("drawing_column_lines"),
                extraction_regions=template_data.get("extraction_regions"),
                extraction_column_lines=template_data.get("extraction_column_lines")
            )
            
            return template_id
            
        except Exception as e:
            print(f"Error saving template: {e}")
            return None
    
    def get_template_safe(self, template_id: int = None, template_name: str = None) -> Optional[Dict]:
        """Safely get a template with error handling"""
        try:
            return self.db.get_template(template_id=template_id, template_name=template_name)
        except Exception as e:
            print(f"Error getting template: {e}")
            return None
    
    def delete_template_safe(self, template_id: int) -> bool:
        """Safely delete a template with error handling"""
        try:
            return self.db.delete_template(template_id)
        except Exception as e:
            print(f"Error deleting template: {e}")
            return False


class UIMessageFactory:
    """Factory for standardized UI messages and dialogs"""
    
    @staticmethod
    def show_error(parent: QWidget, title: str, message: str):
        """Show standardized error message"""
        QMessageBox.critical(parent, title, message, QMessageBox.Ok)
    
    @staticmethod
    def show_warning(parent: QWidget, title: str, message: str):
        """Show standardized warning message"""
        QMessageBox.warning(parent, title, message, QMessageBox.Ok)
    
    @staticmethod
    def show_info(parent: QWidget, title: str, message: str):
        """Show standardized info message"""
        QMessageBox.information(parent, title, message, QMessageBox.Ok)
    
    @staticmethod
    def show_question(parent: QWidget, title: str, message: str) -> bool:
        """Show standardized question dialog"""
        result = QMessageBox.question(
            parent, title, message,
            QMessageBox.Yes | QMessageBox.No
        )
        return result == QMessageBox.Yes
    
    @staticmethod
    def template_exists_error(parent: QWidget, name: str):
        """Show template already exists error"""
        UIMessageFactory.show_warning(
            parent,
            "Template Already Exists",
            f"A template with the name '{name}' already exists. Please choose a different name."
        )
    
    @staticmethod
    def template_saved_success(parent: QWidget, name: str):
        """Show template saved successfully message"""
        UIMessageFactory.show_info(
            parent,
            "Success",
            f"Template '{name}' saved successfully!"
        )
    
    @staticmethod
    def invalid_name_error(parent: QWidget):
        """Show invalid name error"""
        UIMessageFactory.show_warning(
            parent,
            "Invalid Name",
            "Please provide a valid template name."
        )


class ValidationFactory:
    """Factory for common validation operations"""
    
    @staticmethod
    def validate_template_name(name: str) -> bool:
        """Validate template name"""
        return bool(name and name.strip())
    
    @staticmethod
    def validate_template_data(template_data: Dict) -> bool:
        """Validate template data structure"""
        required_fields = ["name", "regions", "column_lines", "config"]
        return all(field in template_data for field in required_fields)
    
    @staticmethod
    def sanitize_template_name(name: str) -> str:
        """Sanitize template name"""
        return name.strip() if name else ""


# Global factory instances for easy access
template_factory = TemplateFactory()
ui_message_factory = UIMessageFactory()
validation_factory = ValidationFactory()

def get_database_factory(db: InvoiceDatabase) -> DatabaseOperationFactory:
    """Get database operation factory instance"""
    return DatabaseOperationFactory(db)
