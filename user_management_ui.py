from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QFormLayout, 
    QLineEdit, QComboBox, QMessageBox, QCheckBox, QGridLayout
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from user_management import UserManagement

class UserManagementDialog(QDialog):
    """Dialog for managing users in the application."""
    
    def __init__(self, user_management, parent=None):
        super().__init__(parent)
        self.user_management = user_management
        self.setup_ui()
        self.load_users()
    
    def setup_ui(self):
        """Set up the user management dialog UI."""
        self.setWindowTitle("User Management")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout(self)
        
        # Header
        header_label = QLabel("User Management")
        header_label.setFont(QFont("Arial", 16))
        layout.addWidget(header_label)
        
        # Description
        description = QLabel("Manage user accounts and permissions.")
        description.setStyleSheet("color: #666;")
        layout.addWidget(description)
        
        # User table
        self.user_table = QTableWidget()
        self.user_table.setColumnCount(5)
        self.user_table.setHorizontalHeaderLabels(["Username", "Full Name", "Email", "Role", "Last Login"])
        self.user_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.user_table)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        self.add_user_button = QPushButton("Add User")
        self.add_user_button.setStyleSheet(
            "background-color: #4169E1; color: white; padding: 8px 16px;"
        )
        self.add_user_button.clicked.connect(self.show_add_user_dialog)
        
        self.edit_user_button = QPushButton("Edit User")
        self.edit_user_button.clicked.connect(self.show_edit_user_dialog)
        
        self.delete_user_button = QPushButton("Delete User")
        self.delete_user_button.setStyleSheet(
            "background-color: #DC3545; color: white; padding: 8px 16px;"
        )
        self.delete_user_button.clicked.connect(self.delete_user)
        
        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)
        
        button_layout.addWidget(self.add_user_button)
        button_layout.addWidget(self.edit_user_button)
        button_layout.addWidget(self.delete_user_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)
    
    def load_users(self):
        """Load all users into the table."""
        try:
            # Clear existing table
            self.user_table.setRowCount(0)
            
            # Query all users
            users = self.user_management.get_all_users()
            
            if not users:
                return
                
            # Fill table with user data
            self.user_table.setRowCount(len(users))
            for row, user in enumerate(users):
                username_item = QTableWidgetItem(user['username'])
                username_item.setData(Qt.UserRole, user['id'])  # Store user ID for later
                self.user_table.setItem(row, 0, username_item)
                
                self.user_table.setItem(row, 1, QTableWidgetItem(user['full_name'] or ""))
                self.user_table.setItem(row, 2, QTableWidgetItem(user['email'] or ""))
                
                # Get role name
                role_name = "N/A"
                if user['role_id']:
                    roles = self.user_management.get_roles()
                    for role in roles:
                        if role['id'] == user['role_id']:
                            role_name = role['name'].capitalize()
                            break
                
                self.user_table.setItem(row, 3, QTableWidgetItem(role_name))
                self.user_table.setItem(row, 4, QTableWidgetItem(user['last_login'] or "Never"))
        
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Error", 
                f"An error occurred while loading users: {str(e)}"
            )
    
    def show_add_user_dialog(self):
        """Show dialog for adding a new user."""
        dialog = UserDialog(self.user_management, parent=self)
        if dialog.exec():
            self.load_users()  # Refresh the list
    
    def show_edit_user_dialog(self):
        """Show dialog for editing an existing user."""
        # Get the selected row
        selected_items = self.user_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(
                self, 
                "Warning", 
                "Please select a user to edit."
            )
            return
        
        # Get the user ID
        row = selected_items[0].row()
        user_id = self.user_table.item(row, 0).data(Qt.UserRole)
        
        # Get user data
        user = self.user_management.get_user_by_id(user_id)
        if not user:
            QMessageBox.warning(
                self, 
                "Warning", 
                "User not found."
            )
            return
        
        # Show edit dialog
        dialog = UserDialog(self.user_management, user=user, parent=self)
        if dialog.exec():
            self.load_users()  # Refresh the list
    
    def delete_user(self):
        """Delete the selected user."""
        # Get the selected row
        selected_items = self.user_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(
                self, 
                "Warning", 
                "Please select a user to delete."
            )
            return
        
        # Get the user ID and username
        row = selected_items[0].row()
        user_id = self.user_table.item(row, 0).data(Qt.UserRole)
        username = self.user_table.item(row, 0).text()
        
        # Confirm deletion
        confirm = QMessageBox.question(
            self, 
            "Confirm Deletion", 
            f"Are you sure you want to delete user '{username}'?\n\nThis action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.Yes:
            # Check if user is currently logged in
            current_user = self.user_management.get_current_user()
            if current_user and current_user['id'] == user_id:
                QMessageBox.warning(
                    self, 
                    "Warning", 
                    "You cannot delete your own account while logged in."
                )
                return
            
            # Delete the user
            if self.user_management.delete_user(user_id):
                QMessageBox.information(
                    self, 
                    "Success", 
                    f"User '{username}' has been deleted."
                )
                self.load_users()  # Refresh the list
            else:
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"An error occurred while deleting the user."
                )


class UserDialog(QDialog):
    """Dialog for adding or editing a user."""
    
    def __init__(self, user_management, user=None, parent=None):
        super().__init__(parent)
        self.user_management = user_management
        self.user = user  # If provided, we're editing an existing user
        self.setup_ui()
        
        # If editing, fill the form with user data
        if user:
            self.fill_form()
    
    def setup_ui(self):
        """Set up the user dialog UI."""
        self.setWindowTitle("Add User" if not self.user else "Edit User")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Form layout
        form_layout = QFormLayout()
        
        # Username field
        self.username_edit = QLineEdit()
        form_layout.addRow("Username:", self.username_edit)
        
        # Password field (only required for new users)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Password:", self.password_edit)
        
        # Confirm password field
        self.confirm_password_edit = QLineEdit()
        self.confirm_password_edit.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Confirm Password:", self.confirm_password_edit)
        
        # Full name field
        self.full_name_edit = QLineEdit()
        form_layout.addRow("Full Name:", self.full_name_edit)
        
        # Email field
        self.email_edit = QLineEdit()
        form_layout.addRow("Email:", self.email_edit)
        
        # Role selection
        self.role_combo = QComboBox()
        self.populate_roles()
        form_layout.addRow("Role:", self.role_combo)
        
        layout.addLayout(form_layout)
        
        # Permissions section (if editing)
        if self.user:
            layout.addSpacing(10)
            
            # Password change notice
            password_notice = QLabel("Leave password fields blank to keep the current password.")
            password_notice.setStyleSheet("color: #666; font-style: italic;")
            layout.addWidget(password_notice)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("Save")
        self.save_button.setStyleSheet(
            "background-color: #4169E1; color: white; padding: 8px 16px;"
        )
        self.save_button.clicked.connect(self.save_user)
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.save_button)
        
        layout.addLayout(button_layout)
    
    def populate_roles(self):
        """Populate the roles combo box."""
        roles = self.user_management.get_roles()
        for role in roles:
            self.role_combo.addItem(role['name'].capitalize(), role['id'])
    
    def fill_form(self):
        """Fill the form with user data for editing."""
        if not self.user:
            return
            
        self.username_edit.setText(self.user['username'])
        self.username_edit.setEnabled(False)  # Cannot change username
        
        self.full_name_edit.setText(self.user['full_name'] or "")
        self.email_edit.setText(self.user['email'] or "")
        
        # Set the role
        if self.user['role_id']:
            role_idx = self.role_combo.findData(self.user['role_id'])
            if role_idx != -1:
                self.role_combo.setCurrentIndex(role_idx)
    
    def save_user(self):
        """Save the user data."""
        # Get the form data
        username = self.username_edit.text().strip()
        password = self.password_edit.text()
        confirm_password = self.confirm_password_edit.text()
        full_name = self.full_name_edit.text().strip()
        email = self.email_edit.text().strip()
        role_id = self.role_combo.currentData()
        
        # Validate the data
        if not username:
            QMessageBox.warning(self, "Warning", "Username is required.")
            return
        
        if not self.user and not password:
            QMessageBox.warning(self, "Warning", "Password is required for new users.")
            return
        
        if password and password != confirm_password:
            QMessageBox.warning(self, "Warning", "Passwords do not match.")
            return
        
        try:
            if self.user:
                # Update existing user
                if password:
                    # Update with new password
                    success = self.user_management.update_user(
                        self.user['id'], full_name, email, role_id, new_password=password
                    )
                else:
                    # Update without changing password
                    success = self.user_management.update_user(
                        self.user['id'], full_name, email, role_id
                    )
                
                if success:
                    QMessageBox.information(self, "Success", f"User '{username}' has been updated.")
                    self.accept()
                else:
                    QMessageBox.critical(self, "Error", "Failed to update user.")
            else:
                # Create new user
                if self.user_management.create_user(username, password, email, full_name, role_id):
                    QMessageBox.information(self, "Success", f"User '{username}' has been created.")
                    self.accept()
                else:
                    QMessageBox.critical(self, "Error", "Failed to create user.")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")


class RoleManagementDialog(QDialog):
    """Dialog for managing roles in the application."""
    
    def __init__(self, user_management, parent=None):
        super().__init__(parent)
        self.user_management = user_management
        self.setup_ui()
        self.load_roles()
    
    def setup_ui(self):
        """Set up the role management dialog UI."""
        self.setWindowTitle("Role Management")
        self.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(self)
        
        # Header
        header_label = QLabel("Role Management")
        header_label.setFont(QFont("Arial", 16))
        layout.addWidget(header_label)
        
        # Description
        description = QLabel("Manage roles and permissions.")
        description.setStyleSheet("color: #666;")
        layout.addWidget(description)
        
        # Roles table
        self.role_table = QTableWidget()
        self.role_table.setColumnCount(3)
        self.role_table.setHorizontalHeaderLabels(["Role Name", "Description", "Permissions"])
        self.role_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.role_table)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        self.add_role_button = QPushButton("Add Role")
        self.add_role_button.setStyleSheet(
            "background-color: #4169E1; color: white; padding: 8px 16px;"
        )
        self.add_role_button.clicked.connect(self.show_add_role_dialog)
        
        self.edit_role_button = QPushButton("Edit Role")
        self.edit_role_button.clicked.connect(self.show_edit_role_dialog)
        
        self.delete_role_button = QPushButton("Delete Role")
        self.delete_role_button.setStyleSheet(
            "background-color: #DC3545; color: white; padding: 8px 16px;"
        )
        self.delete_role_button.clicked.connect(self.delete_role)
        
        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)
        
        button_layout.addWidget(self.add_role_button)
        button_layout.addWidget(self.edit_role_button)
        button_layout.addWidget(self.delete_role_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)
    
    def load_roles(self):
        """Load all roles into the table."""
        try:
            # Clear existing table
            self.role_table.setRowCount(0)
            
            # Get roles
            roles = self.user_management.get_roles()
            
            if not roles:
                return
                
            # Fill table with role data
            self.role_table.setRowCount(len(roles))
            for row, role in enumerate(roles):
                name_item = QTableWidgetItem(role['name'].capitalize())
                name_item.setData(Qt.UserRole, role['id'])  # Store role ID for later
                self.role_table.setItem(row, 0, name_item)
                
                self.role_table.setItem(row, 1, QTableWidgetItem(role['description'] or ""))
                
                # Format permissions as a readable list
                permissions = ", ".join([p.replace("_", " ").capitalize() for p in role['permissions']])
                self.role_table.setItem(row, 2, QTableWidgetItem(permissions))
        
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Error", 
                f"An error occurred while loading roles: {str(e)}"
            )
    
    def show_add_role_dialog(self):
        """Show dialog for adding a new role."""
        dialog = RoleDialog(self.user_management, parent=self)
        if dialog.exec():
            self.load_roles()  # Refresh the list
    
    def show_edit_role_dialog(self):
        """Show dialog for editing an existing role."""
        # Get the selected row
        selected_items = self.role_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(
                self, 
                "Warning", 
                "Please select a role to edit."
            )
            return
        
        # Get the role ID
        row = selected_items[0].row()
        role_id = self.role_table.item(row, 0).data(Qt.UserRole)
        
        # Get role data
        roles = self.user_management.get_roles()
        role = None
        for r in roles:
            if r['id'] == role_id:
                role = r
                break
        
        if not role:
            QMessageBox.warning(
                self, 
                "Warning", 
                "Role not found."
            )
            return
        
        # Show edit dialog
        dialog = RoleDialog(self.user_management, role=role, parent=self)
        if dialog.exec():
            self.load_roles()  # Refresh the list
    
    def delete_role(self):
        """Delete the selected role."""
        # Get the selected row
        selected_items = self.role_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(
                self, 
                "Warning", 
                "Please select a role to delete."
            )
            return
        
        # Get the role ID and name
        row = selected_items[0].row()
        role_id = self.role_table.item(row, 0).data(Qt.UserRole)
        role_name = self.role_table.item(row, 0).text()
        
        # Confirm deletion
        confirm = QMessageBox.question(
            self, 
            "Confirm Deletion", 
            f"Are you sure you want to delete role '{role_name}'?\n\nThis will affect all users with this role.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if confirm == QMessageBox.Yes:
            # Delete the role
            if self.user_management.delete_role(role_id):
                QMessageBox.information(
                    self, 
                    "Success", 
                    f"Role '{role_name}' has been deleted."
                )
                self.load_roles()  # Refresh the list
            else:
                QMessageBox.critical(
                    self, 
                    "Error", 
                    f"An error occurred while deleting the role."
                )


class RoleDialog(QDialog):
    """Dialog for adding or editing a role."""
    
    def __init__(self, user_management, role=None, parent=None):
        super().__init__(parent)
        self.user_management = user_management
        self.role = role  # If provided, we're editing an existing role
        self.setup_ui()
        
        # If editing, fill the form with role data
        if role:
            self.fill_form()
    
    def setup_ui(self):
        """Set up the role dialog UI."""
        self.setWindowTitle("Add Role" if not self.role else "Edit Role")
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout(self)
        
        # Form layout for basic info
        form_layout = QFormLayout()
        
        # Role name field
        self.name_edit = QLineEdit()
        form_layout.addRow("Role Name:", self.name_edit)
        
        # Description field
        self.description_edit = QLineEdit()
        form_layout.addRow("Description:", self.description_edit)
        
        layout.addLayout(form_layout)
        
        # Permissions section
        permissions_group = QLabel("Permissions:")
        layout.addWidget(permissions_group)
        
        # Grid for permissions checkboxes
        permissions_layout = QGridLayout()
        
        # Define available permissions
        self.permissions = {
            "template_management": "Manage Templates",
            "rules_management": "Manage Rules",
            "bulk_extraction": "Bulk Extraction",
            "user_management": "Manage Users",
            "draw_pdf_rules": "Draw PDF Rules"
        }
        
        # Create checkboxes for permissions
        self.permission_checkboxes = {}
        row, col = 0, 0
        for perm_key, perm_label in self.permissions.items():
            checkbox = QCheckBox(perm_label)
            self.permission_checkboxes[perm_key] = checkbox
            permissions_layout.addWidget(checkbox, row, col)
            col += 1
            if col > 1:  # Two columns of checkboxes
                col = 0
                row += 1
        
        layout.addLayout(permissions_layout)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("Save")
        self.save_button.setStyleSheet(
            "background-color: #4169E1; color: white; padding: 8px 16px;"
        )
        self.save_button.clicked.connect(self.save_role)
        
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.save_button)
        
        layout.addLayout(button_layout)
    
    def fill_form(self):
        """Fill the form with role data for editing."""
        if not self.role:
            return
            
        self.name_edit.setText(self.role['name'])
        if self.role['id'] in (1, 2):  # Protect default roles
            self.name_edit.setEnabled(False)
            
        self.description_edit.setText(self.role['description'] or "")
        
        # Set the permissions
        for perm in self.role['permissions']:
            if perm in self.permission_checkboxes:
                self.permission_checkboxes[perm].setChecked(True)
    
    def save_role(self):
        """Save the role data."""
        # Get the form data
        name = self.name_edit.text().strip().lower()
        description = self.description_edit.text().strip()
        
        # Get selected permissions
        permissions = []
        for perm_key, checkbox in self.permission_checkboxes.items():
            if checkbox.isChecked():
                permissions.append(perm_key)
        
        # Validate the data
        if not name:
            QMessageBox.warning(self, "Warning", "Role name is required.")
            return
        
        if not permissions:
            QMessageBox.warning(self, "Warning", "At least one permission must be selected.")
            return
        
        # Format permissions as comma-separated string
        permissions_str = ",".join(permissions)
        
        try:
            if self.role:
                # Update existing role
                if self.user_management.update_role(self.role['id'], name, description, permissions_str):
                    QMessageBox.information(self, "Success", f"Role '{name}' has been updated.")
                    self.accept()
                else:
                    QMessageBox.critical(self, "Error", "Failed to update role.")
            else:
                # Create new role
                if self.user_management.create_role(name, description, permissions_str):
                    QMessageBox.information(self, "Success", f"Role '{name}' has been created.")
                    self.accept()
                else:
                    QMessageBox.critical(self, "Error", "Failed to create role.")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}") 