from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QLineEdit, QPushButton, QMessageBox, QApplication,
                               QTextEdit, QFileDialog, QCheckBox)
from PySide6.QtCore import Qt
import sys
import os
import datetime
from license_manager import get_license_manager

class ActivationDialog(QDialog):
    """
    Dialog for activating the software with a license key
    """
    def __init__(self, parent=None, is_admin=False):
        super().__init__(parent)
        self.license_manager = get_license_manager()
        self.is_admin = is_admin
        self.setWindowTitle("Software Activation")
        self.setMinimumWidth(500)
        self.setMinimumHeight(300)
        self.setup_ui()

    def setup_ui(self):
        """Set up the dialog UI"""
        layout = QVBoxLayout(self)

        # Hardware ID display
        hw_layout = QHBoxLayout()
        hw_label = QLabel("Hardware ID:")
        hw_layout.addWidget(hw_label)

        hw_id = QLineEdit(self.license_manager.hardware_id)
        hw_id.setReadOnly(True)
        hw_layout.addWidget(hw_id)

        copy_btn = QPushButton("Copy")
        copy_btn.clicked.connect(lambda: QApplication.clipboard().setText(self.license_manager.hardware_id))
        hw_layout.addWidget(copy_btn)

        layout.addLayout(hw_layout)

        # License key input
        key_layout = QHBoxLayout()
        key_label = QLabel("License Key:")
        key_layout.addWidget(key_label)

        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("Enter your license key here")
        key_layout.addWidget(self.key_input)

        layout.addLayout(key_layout)

        # Buttons
        btn_layout = QHBoxLayout()

        activate_btn = QPushButton("Activate")
        activate_btn.clicked.connect(self.activate_license)
        btn_layout.addWidget(activate_btn)

        request_btn = QPushButton("Generate License Request")
        request_btn.clicked.connect(self.generate_request)
        btn_layout.addWidget(request_btn)

        # Add unlicense button for admin users
        if self.is_admin:
            unlicense_btn = QPushButton("Remove License")
            unlicense_btn.setStyleSheet("background-color: #e74c3c; color: white;")
            unlicense_btn.clicked.connect(self.remove_license)
            btn_layout.addWidget(unlicense_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

        # Status area
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setPlaceholderText("Activation status will appear here")
        layout.addWidget(self.status_text)

        # Check current license status
        self.check_license()

    def check_license(self):
        """Check the current license status"""
        is_valid, message = self.license_manager.verify_license()
        if is_valid:
            self.status_text.setHtml(f"<span style='color:green'>{message}</span>")
            self.status_text.append("<span style='color:green'>Software is activated and ready to use.</span>")

            # Get more detailed license information
            license_info = self.license_manager.get_license_info()
            if license_info:
                self.status_text.append("<br><span style='color:black'><b>License Details:</b></span>")

                # Edition
                if "edition" in license_info:
                    self.status_text.append(f"<span style='color:black'>Edition: <b>{license_info['edition']}</b></span>")

                # Expiration
                if "expiry_date" in license_info:
                    try:
                        expiry_date = datetime.datetime.fromisoformat(license_info["expiry_date"])
                        days_remaining = license_info.get("days_remaining", 0)

                        if days_remaining > 30:
                            color = "green"
                        elif days_remaining > 0:
                            color = "orange"
                        else:
                            color = "red"

                        self.status_text.append(f"<span style='color:black'>Expiration: <span style='color:{color}'>{expiry_date.strftime('%Y-%m-%d')} ({days_remaining} days remaining)</span></span>")
                    except (ValueError, TypeError):
                        self.status_text.append(f"<span style='color:black'>Expiration: {license_info['expiry_date']}</span>")

                # File limit
                if "file_limit" in license_info:
                    file_limit = license_info["file_limit"]
                    if file_limit == -1:
                        self.status_text.append(f"<span style='color:black'>Bulk Processing: <b>Unlimited files</b></span>")
                    else:
                        self.status_text.append(f"<span style='color:black'>Bulk Processing: <b>Up to {file_limit} files</b></span>")

                # Features
                if "features" in license_info and license_info["features"]:
                    self.status_text.append(f"<span style='color:black'>Features: <b>{', '.join(license_info['features'])}</b></span>")
        else:
            self.status_text.setHtml(f"<span style='color:red'>{message}</span>")
            self.status_text.append("<span style='color:black'>Please enter your license key to activate the software.</span>")
            self.status_text.append("<span style='color:black'>If you don't have a license key, click 'Generate License Request'.</span>")

    def activate_license(self):
        """Activate the software with the entered license key"""
        license_key = self.key_input.text().strip()
        if not license_key:
            QMessageBox.warning(self, "Activation Error", "Please enter a license key.")
            return

        # Try to activate
        success, message = self.license_manager.activate_with_key(license_key)
        if success:
            # Verify the activation
            is_valid, verify_message = self.license_manager.verify_license()
            if is_valid:
                self.status_text.setHtml(f"<span style='color:green'>{message}</span>")
                self.status_text.append("<span style='color:green'>Software activated successfully!</span>")

                QMessageBox.information(self, "Activation Successful",
                                       f"{message}\n\nThe software has been activated successfully.")

                # Update the license information display
                self.check_license()

                self.accept()  # Close dialog with success
            else:
                self.status_text.setHtml(f"<span style='color:red'>{verify_message}</span>")
                self.status_text.append("<span style='color:red'>Activation failed. Please try again.</span>")
        else:
            self.status_text.setHtml(f"<span style='color:red'>{message}</span>")
            self.status_text.append("<span style='color:red'>Please check your license key and try again.</span>")

    def generate_request(self):
        """Generate a license request file"""
        try:
            # Ask where to save the request file
            filename, _ = QFileDialog.getSaveFileName(
                self, "Save License Request", "license_request.json", "JSON Files (*.json)"
            )

            if filename:
                # Generate and save the request
                self.license_manager.save_license_request(filename)

                self.status_text.setHtml("<span style='color:green'>License request generated successfully.</span>")
                self.status_text.append(f"<span style='color:black'>Request saved to: {filename}</span>")
                self.status_text.append("<span style='color:black'>Please send this file to your software vendor to obtain a license key.</span>")

                QMessageBox.information(self, "Request Generated",
                                       f"License request has been saved to:\n{filename}\n\nPlease send this file to your software vendor.")
        except Exception as e:
            self.status_text.setHtml(f"<span style='color:red'>Error generating license request: {str(e)}</span>")
            QMessageBox.warning(self, "Error", f"Failed to generate license request: {str(e)}")

    def remove_license(self):
        """Remove the current license (admin only)"""
        if not self.is_admin:
            QMessageBox.warning(self, "Permission Denied", "Only administrators can remove licenses.")
            return

        # Confirm before removing
        confirm = QMessageBox.question(
            self,
            "Confirm License Removal",
            "Are you sure you want to remove the current license?\n\n"
            "This will deactivate the software on this computer.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if confirm != QMessageBox.Yes:
            return

        # Remove the license
        success, message = self.license_manager.remove_license()

        if success:
            self.status_text.setHtml(f"<span style='color:green'>{message}</span>")
            self.status_text.append("<span style='color:black'>The software has been deactivated on this computer.</span>")
            self.status_text.append("<span style='color:black'>You will need to activate it again to use it.</span>")

            QMessageBox.information(self, "License Removed",
                                   f"{message}\n\nThe software has been deactivated on this computer.")

            # Update the UI to show that no license is active
            self.check_license()
        else:
            self.status_text.setHtml(f"<span style='color:red'>{message}</span>")
            QMessageBox.warning(self, "Error", f"Failed to remove license: {message}")


def check_activation(is_admin=False):
    """
    Check if the software is activated and show activation dialog if needed

    Args:
        is_admin (bool): Whether the current user is an admin (can unlicense)

    Returns:
        bool: True if activated, False otherwise
    """
    license_manager = get_license_manager()
    is_valid, _ = license_manager.verify_license()

    if not is_valid:
        # Show activation dialog
        dialog = ActivationDialog(is_admin=is_admin)
        result = dialog.exec()

        # Check license again after dialog closes
        is_valid, _ = license_manager.verify_license()
        return is_valid

    return True


def show_license_manager_dialog(is_admin=False):
    """
    Show the license manager dialog for managing licenses

    Args:
        is_admin (bool): Whether the current user is an admin (can unlicense)
    """
    dialog = ActivationDialog(is_admin=is_admin)
    dialog.exec()


if __name__ == "__main__":
    # Test the activation dialog
    app = QApplication(sys.argv)
    dialog = ActivationDialog()
    dialog.exec()
    sys.exit(0)
