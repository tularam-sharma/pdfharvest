from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFrame, QDialog, QLineEdit, QFormLayout, QComboBox,
    QMessageBox, QStackedWidget, QGridLayout, QSizePolicy, QScrollArea,
    QGraphicsDropShadowEffect
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QColor
from user_management import UserManagement

class RoleBasedWidget(QWidget):
    """Base class for role-based UI components that only show to users with specific permissions."""

    def __init__(self, required_permission=None, parent=None):
        super().__init__(parent)
        self.required_permission = required_permission
        self.user_management = None

        # Set initial visibility at end of init
        self.initializing = True
        # Call setVisible after we've initialized attributes but don't actually change visibility yet
        self._should_be_visible = not required_permission

    def setVisible(self, visible):
        """Override setVisible to handle initialization sequence properly."""
        if hasattr(self, 'initializing') and self.initializing:
            # During initialization just store the desired state but don't change visibility
            self._should_be_visible = visible
            return

        # Normal case - actually change visibility
        super().setVisible(visible)

    def set_user_management(self, user_management):
        """Set the user management instance and update visibility."""
        self.user_management = user_management
        # Only now apply visibility based on permissions
        self.initializing = False
        self.update_visibility()

    def update_visibility(self):
        """Update the widget visibility based on user permissions."""
        if not self.required_permission:
            self.setVisible(self._should_be_visible)
            return

        if not self.user_management or not self.user_management.get_current_user():
            self.setVisible(False)  # Hide if no user is logged in
            return

        # Show if the user has the required permission
        has_permission = self.user_management.has_permission(self.required_permission)
        print(f"Widget with permission '{self.required_permission}' - has permission: {has_permission}")
        self.setVisible(has_permission)


class UserProfileWidget(RoleBasedWidget):
    """Widget to display the current user's profile and logout option."""
    logout_requested = Signal()
    login_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(required_permission=None, parent=parent)
        self.setup_ui()

    def setup_ui(self):
        """Set up the user profile widget UI."""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)

        self.user_label = QLabel("Not logged in")
        self.user_label.setFont(QFont("Arial", 10))
        layout.addWidget(self.user_label)

        self.role_label = QLabel("")
        self.role_label.setFont(QFont("Arial", 10))
        self.role_label.setStyleSheet("color: #666;")
        layout.addWidget(self.role_label)

        layout.addStretch()

        # # Login button (shown when no user is logged in)
        # self.login_button = QPushButton("Login")
        # self.login_button.setStyleSheet(
        #     "background-color: #4169E1; color: white; padding: 6px 12px;"
        # )
        # self.login_button.clicked.connect(self.login_requested.emit)
        # layout.addWidget(self.login_button)

        # Logout button (shown when a user is logged in)
        self.logout_button = QPushButton("Logout")
        self.logout_button.clicked.connect(self.logout)
        layout.addWidget(self.logout_button)

        # Set initial visibility - will be updated in update_user_info
        # self.login_button.setVisible(True)
        self.logout_button.setVisible(False)

    def update_user_info(self):
        """Update the user information displayed in the widget."""
        if self.user_management and self.user_management.get_current_user():
            user = self.user_management.get_current_user()
            self.user_label.setText(f"Welcome, {user['full_name']}")
            self.role_label.setText(f"Role: {user['role_name'].capitalize()}")
            # self.login_button.setVisible(False)
            self.logout_button.setVisible(True)
        else:
            self.user_label.setText("Not logged in")
            self.role_label.setText("")
            # self.login_button.setVisible(True)
            self.logout_button.setVisible(False)

    def logout(self):
        """Handle user logout."""
        if self.user_management:
            # Ask for confirmation
            confirm = QMessageBox.question(
                self,
                "Confirm Logout",
                "Are you sure you want to log out?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if confirm == QMessageBox.Yes:
                self.user_management.logout()
                self.update_user_info()
                self.logout_requested.emit()

                # Inform the user
                QMessageBox.information(
                    self,
                    "Logged Out",
                    "You have been successfully logged out.",
                    QMessageBox.Ok
                )


class TemplateManagementCard(RoleBasedWidget):
    """Card widget for template management (developers only)."""
    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(required_permission="template_management", parent=parent)
        self.setup_ui()

    def setup_ui(self):
        """Set up the template management card UI."""
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 20px;
                border: 1px solid #ddd;
            }
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 12px;
                border-radius: 6px;
                min-width: 120px;
                text-align: center;
                font-weight: bold;
                min-height: 44px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #3158D3;
            }
            QWidget#buttonContainer {
                background-color: transparent;
                padding: 0;
                border: none;
            }
            QLabel {
                color: #333333;
                font-weight: 500;
            }
            QLabel#titleLabel {
                font-size: 24px;
                font-weight: bold;
                color: #333333;
                margin-bottom: 15px;
            }
        """)
        # Add drop shadow effect
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

        # Set fixed size for cards to prevent stretching
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(300, 500)
        self.setFixedHeight(500)
        # Width will be set dynamically in update_card_layout

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # Icon and title container
        icon_title_container = QWidget()
        icon_title_layout = QVBoxLayout(icon_title_container)
        icon_title_layout.setContentsMargins(0, 0, 0, 0)
        icon_title_layout.setSpacing(10)

        # Icon
        icon_label = QLabel("⚙️")
        icon_label.setFont(QFont("Arial", 36))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_title_layout.addWidget(icon_label)

        # Title
        title_label = QLabel("Template Management")
        title_label.setObjectName("titleLabel")
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setWordWrap(True)
        icon_title_layout.addWidget(title_label)

        layout.addWidget(icon_title_container)

        # Subtitle
        subtitle_label = QLabel("Create and manage invoice templates")
        subtitle_label.setFont(QFont("Arial", 14))
        subtitle_label.setStyleSheet("color: #333;")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setWordWrap(True)
        layout.addWidget(subtitle_label)

        # Description
        description_label = QLabel(
            "Save extraction settings as templates for processing similar invoices. "
        )
        description_label.setWordWrap(True)
        description_label.setFont(QFont("Arial", 12))
        description_label.setStyleSheet("color: #333;")
        description_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(description_label)

        layout.addStretch()

        # Button container
        button_container = QWidget()
        button_container.setObjectName("buttonContainer")
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(20, 0, 20, 0)

        # Action button
        self.action_button = QPushButton("Manage Templates")
        self.action_button.setObjectName("loginButton")
        self.action_button.setMinimumHeight(44)
        self.action_button.setCursor(Qt.PointingHandCursor)
        self.action_button.clicked.connect(self.clicked.emit)
        button_layout.addWidget(self.action_button, alignment=Qt.AlignCenter)

        layout.addWidget(button_container)


# RulesManagementCard removed as it is no longer needed


class BulkExtractionCard(RoleBasedWidget):
    """Card widget for bulk extraction (available to all users)."""
    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(required_permission="bulk_extraction", parent=parent)
        self.setup_ui()

    def setup_ui(self):
        """Set up the bulk extraction card UI."""
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 20px;
                border: 1px solid #ddd;
            }
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 12px;
                border-radius: 6px;
                min-width: 120px;
                text-align: center;
                font-weight: bold;
                min-height: 44px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #3158D3;
            }
            QWidget#buttonContainer {
                background-color: transparent;
                padding: 0;
                border: none;
            }
            QLabel {
                color: #333333;
                font-weight: 500;
            }
            QLabel#titleLabel {
                font-size: 24px;
                font-weight: bold;
                color: #333333;
                margin-bottom: 15px;
            }
        """)
        # Add drop shadow effect
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

        # Set fixed size for cards to prevent stretching
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(300, 500)
        self.setFixedHeight(500)
        # Width will be set dynamically in update_card_layout

        layout = QVBoxLayout(self)
        layout.setSpacing(20)  # Increased spacing between elements
        layout.setContentsMargins(20, 30, 20, 30)  # Increased vertical margins

        # Icon and title container
        icon_title_container = QWidget()
        icon_title_layout = QVBoxLayout(icon_title_container)
        icon_title_layout.setContentsMargins(0, 0, 0, 0)
        icon_title_layout.setSpacing(10)

        # Icon
        icon_label = QLabel("📦")
        icon_label.setFont(QFont("Arial", 36))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_title_layout.addWidget(icon_label)

        # Title
        title_label = QLabel("Bulk Extraction")
        title_label.setObjectName("titleLabel")
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setWordWrap(True)
        icon_title_layout.addWidget(title_label)

        layout.addWidget(icon_title_container)

        # Subtitle
        subtitle_label = QLabel("Process multiple invoices at once")
        subtitle_label.setFont(QFont("Arial", 14))
        subtitle_label.setStyleSheet("color: #333;")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setWordWrap(True)
        layout.addWidget(subtitle_label)

        # Description
        description_label = QLabel(
            "Upload multiple PDF invoices and process them in batch using saved templates. "
            "Export results in various formats."
        )
        description_label.setWordWrap(True)
        description_label.setFont(QFont("Arial", 12))
        description_label.setStyleSheet("color: #333;")
        description_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(description_label)

        # Add more stretch to push content apart vertically
        layout.addStretch(3)

        # Button container
        button_container = QWidget()
        button_container.setObjectName("buttonContainer")
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(10, 0, 10, 0)  # Reduced horizontal margins

        # Action button
        self.action_button = QPushButton("Bulk Process")
        self.action_button.setObjectName("loginButton")
        self.action_button.setMinimumHeight(44)
        self.action_button.setMinimumWidth(250)  # Even wider minimum width
        self.action_button.setCursor(Qt.PointingHandCursor)
        self.action_button.clicked.connect(self.clicked.emit)
        button_layout.addWidget(self.action_button, alignment=Qt.AlignCenter)

        layout.addWidget(button_container)


class UploadProcessCard(RoleBasedWidget):
    """Card widget for single invoice upload and processing."""
    clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(required_permission="draw_pdf_rules", parent=parent)  # Use draw_pdf_rules permission
        self.setup_ui()

    def setup_ui(self):
        """Set up the upload and process card UI."""
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 20px;
                border: 1px solid #ddd;
            }
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 12px;
                border-radius: 6px;
                min-width: 120px;
                text-align: center;
                font-weight: bold;
                min-height: 44px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #3158D3;
            }
            QWidget#buttonContainer {
                background-color: transparent;
                padding: 0;
                border: none;
            }
            QLabel {
                color: #333333;
                font-weight: 500;
            }
            QLabel#titleLabel {
                font-size: 24px;
                font-weight: bold;
                color: #333333;
                margin-bottom: 15px;
            }
        """)
        # Add drop shadow effect
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

        # Set fixed size for cards to prevent stretching
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setMinimumSize(300, 500)
        self.setFixedHeight(500)
        # Width will be set dynamically in update_card_layout

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # Icon and title container
        icon_title_container = QWidget()
        icon_title_layout = QVBoxLayout(icon_title_container)
        icon_title_layout.setContentsMargins(0, 0, 0, 0)
        icon_title_layout.setSpacing(10)

        # Icon
        icon_label = QLabel("⬆️")
        icon_label.setFont(QFont("Arial", 36))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_title_layout.addWidget(icon_label)

        # Title
        title_label = QLabel("Upload & Process")
        title_label.setObjectName("titleLabel")
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setWordWrap(True)
        icon_title_layout.addWidget(title_label)

        layout.addWidget(icon_title_container)

        # Subtitle
        subtitle_label = QLabel("Upload a PDF invoice and extract tables visually")
        subtitle_label.setFont(QFont("Arial", 14))
        subtitle_label.setStyleSheet("color: #333;")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setWordWrap(True)
        layout.addWidget(subtitle_label)

        # Description
        description_label = QLabel(
            "Upload your invoice PDF, select table regions visually, and configure column mapping "
            "to extract structured data."
        )
        description_label.setWordWrap(True)
        description_label.setFont(QFont("Arial", 12))
        description_label.setStyleSheet("color: #333;")
        description_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(description_label)

        layout.addStretch()

        # Button container
        button_container = QWidget()
        button_container.setObjectName("buttonContainer")
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(20, 0, 20, 0)

        # Action button
        self.action_button = QPushButton("Get Started")
        self.action_button.setObjectName("loginButton")
        self.action_button.setMinimumHeight(44)
        self.action_button.setCursor(Qt.PointingHandCursor)
        self.action_button.clicked.connect(self.clicked.emit)
        button_layout.addWidget(self.action_button, alignment=Qt.AlignCenter)

        layout.addWidget(button_container)


class MainDashboard(QWidget):
    """Main dashboard with role-based cards."""
    show_pdf_processor = Signal()
    show_template_manager = Signal()
    # show_rules_manager signal removed
    show_bulk_processor = Signal()
    show_user_management = Signal()  # New signal for user management
    login_successful = Signal(dict)  # Signal for successful login

    def __init__(self, user_management, parent=None):
        super().__init__(parent)
        self.user_management = user_management
        self.setup_ui()

        # Connect to resize events to update card layout
        self.resizeEvent = self.on_resize

    def on_resize(self, event):
        """Handle resize events to update layouts."""
        self.update_card_layout()

        # Also adjust login panel size on resize
        if hasattr(self, 'login_panel') and self.login_panel.isVisible():
            self.adjust_login_panel_size()

        # Call the parent implementation
        super().resizeEvent(event)

    def setup_ui(self):
        """Set up the main dashboard UI."""
        main_layout = QVBoxLayout(self)

        # User profile - we'll keep this for logged-in state
        self.user_profile = UserProfileWidget()
        self.user_profile.set_user_management(self.user_management)
        self.user_profile.update_user_info()
        self.user_profile.logout_requested.connect(self.handle_logout)
        main_layout.addWidget(self.user_profile)

        # Create header
        header_layout = QHBoxLayout()
        logo_label = QLabel("📄 Invoice Harvest")
        logo_label.setFont(QFont("Arial", 16))
        header_layout.addWidget(logo_label)

        # Add header buttons - these will be shown/hidden based on permissions
        header_layout.addStretch()

        # User Management button
        self.user_management_button = QPushButton("User Management")
        self.user_management_button.setStyleSheet(
            "background-color: #4169E1; color: white; padding: 6px 12px;"
        )
        self.user_management_button.clicked.connect(self.show_user_management.emit)
        self.user_management_button.setVisible(False)  # Hidden by default
        header_layout.addWidget(self.user_management_button)

        main_layout.addLayout(header_layout)

        # Create content area with scroll capabilities for small screens
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)  # Remove border

        # Create content widget for scroll area
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(20, 20, 20, 20)

        # # Add title and subtitle
        # title = QLabel("")
        # title.setFont(QFont("Arial", 32))
        # title.setAlignment(Qt.AlignCenter)
        # subtitle = QLabel("Easily extract tables from PDF invoices with visual selection and mapping")
        # subtitle.setFont(QFont("Arial", 14))
        # subtitle.setAlignment(Qt.AlignCenter)
        # subtitle.setStyleSheet("color: #666;")
        # subtitle.setWordWrap(True)

        # self.content_layout.addWidget(title)
        # self.content_layout.addWidget(subtitle)
        # self.content_layout.addSpacing(20)

        # Create login panel (will be shown/hidden based on login state)
        # Create a QFrame to contain the login panel with border styling
        self.login_container_frame = QFrame()
        self.login_container_frame.setFrameShape(QFrame.StyledPanel)
        self.login_container_frame.setFrameShadow(QFrame.Raised)
        self.login_container_frame.setObjectName("loginContainerFrame")
        self.login_container_frame.setStyleSheet("""
            #loginContainerFrame {
                background-color: transparent;
                border: none;
            }
        """)

        # Create the actual login panel as a QWidget inside the frame
        self.login_panel = QWidget(self.login_container_frame)
        self.login_panel.setObjectName("loginPanel")
        self.login_panel.setStyleSheet("""
            #loginPanel {
                background-color: white;
                border-radius: 10px;
                min-height: 350px;
                max-width: 400px;
                min-width: 280px;
                margin: 0 auto;
                border: 1px solid #ddd;
            }
            QLabel {
                font-size: 14px;
                color: #333333;
                font-weight: 500;
            }
            QLineEdit {
                padding: 10px 12px;
                border: 1px solid #ddd;
                border-radius: 6px;
                font-size: 14px;
                background-color: #f8f9fa;
                color: #333333;
                min-height: 20px;
                selection-background-color: #4169E1;
            }
            QLineEdit:focus {
                border: 1px solid #4169E1;
                background-color: white;
            }
            QLineEdit::placeholder {
                color: #aaa;
            }
            QPushButton#loginButton {
                background-color: #4169E1;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px;
                font-size: 16px;
                font-weight: bold;
                min-height: 44px;
            }
            QPushButton#loginButton:hover {
                background-color: #3158D3;
            }
            QPushButton#loginButton:pressed {
                background-color: #2A4CC9;
            }
            QLabel#errorLabel {
                color: #d32f2f;
                font-size: 13px;
                min-height: 20px;
                font-weight: normal;
            }
            QLabel#titleLabel {
                font-size: 24px;
                font-weight: bold;
                color: #333333;
                margin-bottom: 15px;
            }
        """)

        # Set layout for the login panel
        login_layout = QVBoxLayout(self.login_panel)
        login_layout.setContentsMargins(30, 30, 30, 30)
        login_layout.setSpacing(12)

        # Login form title
        login_title = QLabel("Invoice Harvest")
        login_title.setObjectName("titleLabel")
        login_title.setAlignment(Qt.AlignCenter)
        login_layout.addWidget(login_title)
        login_layout.addSpacing(10)

        # Form container with proper margins
        form_container = QWidget()
        form_layout = QVBoxLayout(form_container)
        form_layout.setContentsMargins(30, 30, 30, 30)
        form_layout.setSpacing(8)

        # Username field
        username_label = QLabel("Username")
        form_layout.addWidget(username_label)

        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("Enter your username")
        self.username_edit.setMinimumHeight(44)
        form_layout.addWidget(self.username_edit)
        form_layout.addSpacing(16)

        # Password field
        password_label = QLabel("Password")
        form_layout.addWidget(password_label)

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("Enter your password")
        self.password_edit.setMinimumHeight(44)
        self.password_edit.returnPressed.connect(self.handle_login)  # Handle Enter key
        form_layout.addWidget(self.password_edit)
        form_layout.addSpacing(24)

        # Add form container to main layout
        login_layout.addWidget(form_container)

        # Login button
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(30, 0, 30, 0)

        self.login_button = QPushButton("Sign In")
        self.login_button.setObjectName("loginButton")
        self.login_button.setMinimumHeight(44)
        self.login_button.setMaximumWidth(340)
        self.login_button.setCursor(Qt.PointingHandCursor)  # Add hand cursor
        self.login_button.clicked.connect(self.handle_login)
        button_layout.addWidget(self.login_button)

        login_layout.addWidget(button_container)

        # Error message
        self.error_label = QLabel("")
        self.error_label.setObjectName("errorLabel")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setMinimumHeight(20)
        login_layout.addWidget(self.error_label)

        # Add stretch to push everything up
        login_layout.addStretch()

        # Set layout for the frame
        frame_layout = QVBoxLayout(self.login_container_frame)
        frame_layout.addWidget(self.login_panel)
        frame_layout.setContentsMargins(3, 3, 3, 3)  # Small margins to show shadow

        # Add login panel centered in layout
        login_container = QHBoxLayout()
        login_container.addStretch()
        login_container.addWidget(self.login_container_frame)
        login_container.addStretch()

        # Create cards container widget (will be shown after login)
        self.cards_container = QWidget()
        self.cards_container.setStyleSheet("""
            QWidget {
                background-color: transparent;
            }
        """)
        self.cards_layout = QGridLayout(self.cards_container)
        self.cards_layout.setSpacing(20)  # Space between cards
        self.cards_layout.setContentsMargins(10, 10, 10, 10)  # Reduced margins
        # No alignment setting to allow natural left alignment

        # Upload & Process card - connect to check_permission method first
        self.upload_card = UploadProcessCard()
        self.upload_card.set_user_management(self.user_management)
        # Connect to our permission check method instead of direct signal
        self.upload_card.clicked.connect(self.check_pdf_processor_permission)

        # Template Management card (developers only)
        self.template_card = TemplateManagementCard()
        self.template_card.set_user_management(self.user_management)
        self.template_card.clicked.connect(self.show_template_manager.emit)

        # Rules Management card removed

        # Bulk Extraction card (all users)
        self.bulk_card = BulkExtractionCard()
        self.bulk_card.set_user_management(self.user_management)
        self.bulk_card.clicked.connect(self.show_bulk_processor.emit)

        # Add cards to list - we'll arrange them dynamically
        self.cards = [self.upload_card, self.template_card, self.bulk_card]

        self.content_layout.addLayout(login_container)

        # Add cards container directly to allow natural alignment
        self.content_layout.addWidget(self.cards_container)

        self.content_layout.addStretch()

        # Set the content widget as the scroll area's widget
        self.scroll_area.setWidget(self.content_widget)
        main_layout.addWidget(self.scroll_area)

        # Initialize card visibility based on current permissions
        self.update_dashboard()

    def handle_login(self):
        """Handle login directly from the home page."""
        username = self.username_edit.text()
        password = self.password_edit.text()

        if not username or not password:
            self.error_label.setText("Please enter both username and password")
            return

        user = self.user_management.authenticate_user(username, password)
        if user:
            # Login successful
            self.error_label.setText("")

            # Emit signal for main app to handle (this will trigger handle_login_success in main.py)
            self.login_successful.emit(user)

            # Update dashboard UI
            self.update_dashboard()

            # Clear fields for next login
            self.username_edit.clear()
            self.password_edit.clear()

            # We don't show a message box here as the main app will do that
            # The welcome message is handled by handle_login_success in PDFHarvest class
        else:
            self.error_label.setText("Invalid username or password")

    def handle_logout(self):
        """Handle logout and show login panel."""
        self.update_dashboard()

    def check_pdf_processor_permission(self):
        """Check if user has permission to access PDF processor, show login prompt if not"""
        if not self.user_management.get_current_user() or not self.user_management.has_permission("draw_pdf_rules"):
            # Show permission denied message with login option
            result = QMessageBox.question(
                self,
                "Permission Required",
                "You need to log in with appropriate permissions to access this feature.\n\nWould you like to log in now?",
                QMessageBox.Yes | QMessageBox.No
            )

            if result == QMessageBox.Yes:
                # Focus on the username field of the login panel
                self.login_panel.setVisible(True)
                self.update_dashboard()
                self.username_edit.setFocus()
        else:
            # User has permission, proceed to PDF processor
            self.show_pdf_processor.emit()

    def update_dashboard(self):
        """Update dashboard components based on user permissions."""
        self.user_profile.update_user_info()

        # Toggle visibility of login panel vs cards based on login state
        is_logged_in = bool(self.user_management and self.user_management.get_current_user())

        # Show login panel only if not logged in
        self.login_panel.setVisible(not is_logged_in)

        # Update card visibility - hide cards container when showing login
        self.cards_container.setVisible(is_logged_in)

        # Adjust login panel size based on screen width
        self.adjust_login_panel_size()

        # Update each card's visibility based on permissions
        if is_logged_in:
            self.upload_card.update_visibility()
            self.template_card.update_visibility()
            # rules_card removed
            self.bulk_card.update_visibility()

            # Update card layout based on current visibility
            self.update_card_layout()

        # Update user management button visibility
        has_user_management = False
        if self.user_management and self.user_management.get_current_user():
            has_user_management = self.user_management.has_permission('user_management')
        self.user_management_button.setVisible(has_user_management)

    def adjust_login_panel_size(self):
        """Adjust login panel size based on screen size."""
        content_width = self.content_widget.width()

        # On small screens, login panel should take more width
        if content_width < 600:
            # Smaller screens - panel takes up to 90% of width
            max_width = min(content_width * 0.9, 400)
            min_width = min(content_width * 0.85, 280)

            self.login_panel.setStyleSheet(self.login_panel.styleSheet() + f"""
                #loginPanel {{
                    max-width: {int(max_width)}px;
                    min-width: {int(min_width)}px;
                }}
            """)
        else:
            # Larger screens - standard width
            self.login_panel.setStyleSheet(self.login_panel.styleSheet() + """
                #loginPanel {
                    max-width: 400px;
                    min-width: 350px;
                }
            """)

    def update_card_layout(self):
        """Update card layout based on available width and visible cards."""
        # Clear existing layout
        while self.cards_layout.count():
            item = self.cards_layout.takeAt(0)
            if item:
                widget = item.widget()
                if widget:
                    self.cards_layout.removeWidget(widget)

        # Get visible cards
        visible_cards = [card for card in self.cards if not card.isHidden()]
        if not visible_cards:
            self.cards_container.setVisible(False)
            return  # No visible cards

        # Determine optimal number of columns based on container width
        container_width = self.cards_container.width()

        # Use dynamic column calculation based on width
        # For smaller screens, use 1 column; for medium, use 2; for large, use 3
        if container_width < 600:
            max_cols = 1
        elif container_width < 1000:
            max_cols = 2
        else:
            max_cols = 3

        # Calculate fixed card width based on container width and columns
        # Account for spacing between cards (20px) and container margins
        spacing = self.cards_layout.spacing()
        total_spacing = spacing * (max_cols - 1)
        margins_lr = self.cards_layout.contentsMargins().left() + self.cards_layout.contentsMargins().right()
        card_width = (container_width - total_spacing - margins_lr) / max_cols

        # Set fixed width for all cards
        for card in visible_cards:
            card.setFixedWidth(int(card_width))

        # Calculate how many rows and columns we need
        num_cards = len(visible_cards)
        num_rows = (num_cards + max_cols - 1) // max_cols  # Ceiling division

        # Always use a standard layout regardless of number of cards
        # This ensures consistent positioning even with few cards
        row, col = 0, 0
        for card in visible_cards:
            self.cards_layout.addWidget(card, row, col)
            col += 1
            if col >= max_cols:
                col = 0
                row += 1


class RoleBasedPDFProcessor(RoleBasedWidget):
    """Role-based wrapper for the PDF Processor that requires draw_pdf_rules permission."""

    def __init__(self, parent=None):
        # Initialize attributes that will be set in setup_ui
        self.permission_message = None
        self.pdf_processor = None

        # Require draw_pdf_rules permission for this screen
        super().__init__(required_permission="draw_pdf_rules", parent=parent)

        # Now set up the UI after parent initialization
        self.setup_ui()

    def setup_ui(self):
        """Set up the PDF Processor UI with permission control."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Create the permission denied message
        self.permission_message = QFrame(self)
        self.permission_message.setStyleSheet("""
            QFrame {
                background-color: #fff8e1;
                border: 1px solid #ffd54f;
                border-radius: 8px;
                padding: 16px;
                margin: 32px;
            }
            QLabel {
                color: #ff6f00;
                font-size: 16px;
            }
        """)

        message_layout = QVBoxLayout(self.permission_message)

        # Add warning icon and text
        error_label = QLabel("⚠️ Permission Denied")
        error_label.setFont(QFont("Arial", 18, QFont.Bold))
        error_label.setAlignment(Qt.AlignCenter)

        details_label = QLabel(
            "You do not have permission to access the PDF Processor.\n\n"
            "This feature requires the 'draw_pdf_rules' permission.\n\n"
            "Please contact your administrator or log in with appropriate credentials."
        )
        details_label.setAlignment(Qt.AlignCenter)
        details_label.setWordWrap(True)

        login_button = QPushButton("Log In")
        login_button.setStyleSheet("""
            QPushButton {
                background-color: #4169E1;
                color: white;
                padding: 10px 20px;
                border-radius: 4px;
                min-width: 120px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3158D3;
            }
        """)
        login_button.clicked.connect(self.request_login)

        message_layout.addWidget(error_label)
        message_layout.addWidget(details_label)
        message_layout.addWidget(login_button, 0, Qt.AlignCenter)

        # Create the actual PDF Processor (hidden initially)
        from pdf_processor import PDFProcessor
        self.pdf_processor = PDFProcessor()
        self.pdf_processor.setVisible(False)  # Initially hidden

        # Add both widgets to the layout
        layout.addWidget(self.permission_message)
        layout.addWidget(self.pdf_processor)

        # Connect any signals that need to be passed through
        if hasattr(self.pdf_processor, 'next_clicked'):
            self.pdf_processor.next_clicked.connect(self.next_clicked)

        # Initial setup - update visibility based on permissions
        self.update_widget_visibility()

    def setVisible(self, visible):
        """Override setVisible to ensure proper internal widget visibility."""
        super().setVisible(visible)

        # When the entire widget is hidden, ensure both subwidgets are hidden too
        if not visible:
            if self.permission_message:
                self.permission_message.setVisible(False)
            if self.pdf_processor:
                self.pdf_processor.setVisible(False)
        else:
            # When shown, update internal widget visibility based on permissions
            self.update_widget_visibility()

    def set_user_management(self, user_management):
        """Set the user management instance and update visibility."""
        super().set_user_management(user_management)
        self.update_widget_visibility()

    def update_visibility(self):
        """Update widget visibility when permissions change."""
        super().update_visibility()
        self.update_widget_visibility()

    def update_widget_visibility(self):
        """Update internal widget visibility based on permissions."""
        # Only proceed if both widgets have been initialized
        if not hasattr(self, 'permission_message') or not hasattr(self, 'pdf_processor'):
            return

        if self.permission_message is None or self.pdf_processor is None:
            return

        has_permission = False
        if self.user_management and self.user_management.get_current_user():
            has_permission = self.user_management.has_permission(self.required_permission)

        # Show/hide the appropriate widgets only if this widget is visible
        if self.isVisible():
            self.permission_message.setVisible(not has_permission)
            self.pdf_processor.setVisible(has_permission)

    def request_login(self):
        """Request to show the login panel in the dashboard."""
        # Find the parent window
        main_window = self.window()

        # Get reference to main dashboard and switch to it
        for i in range(main_window.centralWidget().count()):
            widget = main_window.centralWidget().widget(i)
            if isinstance(widget, MainDashboard):
                # Switch to dashboard
                main_window.centralWidget().setCurrentWidget(widget)
                # Show login panel
                widget.update_dashboard()
                # Focus on username field if available
                if hasattr(widget, 'username_edit'):
                    widget.username_edit.setFocus()
                break

    # Forward necessary methods to the PDF processor
    def load_pdf_file(self, filepath, skip_dialog=False):
        """Load a PDF file in the processor."""
        print(f"\n[DEBUG] load_pdf_file called in RoleBasedPDFProcessor")
        print(f"[DEBUG] File path: {filepath}")
        print(f"[DEBUG] Skip dialog: {skip_dialog}")
        print(f"[DEBUG] PDF processor visible: {self.pdf_processor.isVisible()}")

        # Always load the file, regardless of visibility
        self.pdf_processor.load_pdf_file(filepath, skip_dialog=skip_dialog)

        # Make sure the PDF processor is visible
        self.pdf_processor.setVisible(True)
        self.permission_message.setVisible(False)

        print(f"[DEBUG] PDF processor visibility after: {self.pdf_processor.isVisible()}")
        print(f"[DEBUG] Permission message visibility after: {self.permission_message.isVisible()}")

    # Forward properties to the PDF processor
    @property
    def pdf_path(self):
        return self.pdf_processor.pdf_path if hasattr(self.pdf_processor, 'pdf_path') else None

    @pdf_path.setter
    def pdf_path(self, value):
        if hasattr(self.pdf_processor, 'pdf_path'):
            self.pdf_processor.pdf_path = value

    @property
    def regions(self):
        return self.pdf_processor.regions if hasattr(self.pdf_processor, 'regions') else None

    @regions.setter
    def regions(self, value):
        if hasattr(self.pdf_processor, 'regions'):
            self.pdf_processor.regions = value

    @property
    def column_lines(self):
        return self.pdf_processor.column_lines if hasattr(self.pdf_processor, 'column_lines') else None

    @column_lines.setter
    def column_lines(self, value):
        if hasattr(self.pdf_processor, 'column_lines'):
            self.pdf_processor.column_lines = value

    @property
    def config(self):
        return self.pdf_processor.config if hasattr(self.pdf_processor, 'config') else {}

    @config.setter
    def config(self, value):
        if hasattr(self.pdf_processor, 'config'):
            self.pdf_processor.config = value

    @property
    def multi_table_mode(self):
        return self.pdf_processor.multi_table_mode if hasattr(self.pdf_processor, 'multi_table_mode') else False

    @multi_table_mode.setter
    def multi_table_mode(self, value):
        if hasattr(self.pdf_processor, 'multi_table_mode'):
            self.pdf_processor.multi_table_mode = value

    @property
    def multi_page_mode(self):
        return self.pdf_processor.multi_page_mode if hasattr(self.pdf_processor, 'multi_page_mode') else False

    @multi_page_mode.setter
    def multi_page_mode(self, value):
        if hasattr(self.pdf_processor, 'multi_page_mode'):
            self.pdf_processor.multi_page_mode = value

    @property
    def page_regions(self):
        return self.pdf_processor.page_regions if hasattr(self.pdf_processor, 'page_regions') else {}

    @page_regions.setter
    def page_regions(self, value):
        if hasattr(self.pdf_processor, 'page_regions'):
            self.pdf_processor.page_regions = value

    @property
    def page_configs(self):
        return self.pdf_processor.page_configs if hasattr(self.pdf_processor, 'page_configs') else []

    @page_configs.setter
    def page_configs(self, value):
        if hasattr(self.pdf_processor, 'page_configs'):
            self.pdf_processor.page_configs = value

    @property
    def page_column_lines(self):
        return self.pdf_processor.page_column_lines if hasattr(self.pdf_processor, 'page_column_lines') else {}

    @page_column_lines.setter
    def page_column_lines(self, value):
        if hasattr(self.pdf_processor, 'page_column_lines'):
            self.pdf_processor.page_column_lines = value

    @property
    def pdf_label(self):
        """Access to the underlying PDFProcessor's pdf_label"""
        return self.pdf_processor.pdf_label if hasattr(self.pdf_processor, 'pdf_label') else None

    @property
    def use_middle_page(self):
        return self.pdf_processor.use_middle_page if hasattr(self.pdf_processor, 'use_middle_page') else False

    @use_middle_page.setter
    def use_middle_page(self, value):
        if hasattr(self.pdf_processor, 'use_middle_page'):
            self.pdf_processor.use_middle_page = value

    @property
    def fixed_page_count(self):
        return self.pdf_processor.fixed_page_count if hasattr(self.pdf_processor, 'fixed_page_count') else False

    @fixed_page_count.setter
    def fixed_page_count(self, value):
        if hasattr(self.pdf_processor, 'fixed_page_count'):
            self.pdf_processor.fixed_page_count = value

    # Add any signal this widget should emit
    next_clicked = Signal()