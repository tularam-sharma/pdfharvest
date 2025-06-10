import sqlite3
import hashlib
import uuid
import os
from datetime import datetime

# Import path_helper for path resolution (optional)
try:
    import path_helper
    PATH_HELPER_AVAILABLE = True
except ImportError:
    PATH_HELPER_AVAILABLE = False
    print("[WARNING] path_helper not available. Using default path resolution.")

class UserManagement:
    def __init__(self, db_path="user_management.db"):
        """Initialize the user management system with the database."""
        # Resolve the database path
        if PATH_HELPER_AVAILABLE:
            self.db_path = path_helper.resolve_path(db_path)
        else:
            self.db_path = os.path.abspath(db_path)
        self.current_user = None
        self.create_connection()
        self.create_tables()

    def create_connection(self):
        """Create a database connection to the SQLite database."""
        try:
            # Ensure the database file is writable if it exists
            if os.path.exists(self.db_path):
                try:
                    os.chmod(self.db_path, 0o600)  # Read-write for owner
                    print(f"Ensured database file is writable: {self.db_path}")
                except Exception as chmod_e:
                    print(f"Warning: Could not set file permissions: {str(chmod_e)}")

            # Connect to the database with write permissions
            self.conn = sqlite3.connect(self.db_path)
            self.conn.execute('PRAGMA journal_mode=WAL')  # Use Write-Ahead Logging for better concurrency
            self.cursor = self.conn.cursor()

            # Test if we can write to the database
            self.cursor.execute('PRAGMA user_version = 1')
            self.conn.commit()

            print(f"Connected to user management database: {self.db_path}")
        except sqlite3.Error as e:
            print(f"Database connection error: {e}")
            import traceback
            traceback.print_exc()

            # Try to recover by recreating the database if it's read-only
            if 'readonly database' in str(e).lower():
                print("Attempting to recover from read-only database error...")
                try:
                    # Close the current connection if it exists
                    if hasattr(self, 'conn') and self.conn:
                        self.conn.close()

                    # Try to delete the database file and create a new one
                    if os.path.exists(self.db_path):
                        os.remove(self.db_path)
                        print(f"Removed read-only database: {self.db_path}")

                    # Create a new connection
                    self.conn = sqlite3.connect(self.db_path)
                    self.cursor = self.conn.cursor()
                    print(f"Created new database: {self.db_path}")
                except Exception as recovery_e:
                    print(f"Recovery failed: {str(recovery_e)}")
                    traceback.print_exc()

    def create_tables(self):
        """Create the users and roles tables if they don't exist."""
        # Create roles table
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS roles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            description TEXT,
            permissions TEXT NOT NULL
        )
        ''')

        # Create users table
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            email TEXT UNIQUE,
            full_name TEXT,
            role_id INTEGER,
            created_at TEXT NOT NULL,
            last_login TEXT,
            FOREIGN KEY (role_id) REFERENCES roles (id)
        )
        ''')

        # Check if the tables have data
        self.cursor.execute("SELECT COUNT(*) FROM roles")
        roles_count = self.cursor.fetchone()[0]

        self.cursor.execute("SELECT COUNT(*) FROM users")
        users_count = self.cursor.fetchone()[0]

        # If either table is empty, recreate the default data
        if roles_count == 0 or users_count == 0:
            # Clear any existing data if partial
            if roles_count > 0:
                self.cursor.execute("DELETE FROM roles")
            if users_count > 0:
                self.cursor.execute("DELETE FROM users")

            # Add default roles
            self.cursor.execute('''
            INSERT INTO roles (id, name, description, permissions)
            VALUES (?, ?, ?, ?)
            ''', (1, 'developer', 'System developer with full access',
                  'template_management,rules_management,bulk_extraction,user_management,draw_pdf_rules'))

            self.cursor.execute('''
            INSERT INTO roles (id, name, description, permissions)
            VALUES (?, ?, ?, ?)
            ''', (2, 'user', 'Regular user with limited access',
                  'bulk_extraction'))

            # Add default admin user
            admin_password = self._hash_password('admin')
            self.cursor.execute('''
            INSERT INTO users (username, password_hash, email, full_name, role_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', ('admin', admin_password, 'admin@example.com', 'Admin User', 1, datetime.now().isoformat()))

            # Add default regular user
            user_password = self._hash_password('user')
            self.cursor.execute('''
            INSERT INTO users (username, password_hash, email, full_name, role_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', ('user', user_password, 'user@example.com', 'Regular User', 2, datetime.now().isoformat()))

        self.conn.commit()
        print("User management tables created and initialized")

    def _hash_password(self, password):
        """Hash a password for storing."""
        salt = uuid.uuid4().hex
        return hashlib.sha256(salt.encode() + password.encode()).hexdigest() + ':' + salt

    def _verify_password(self, stored_password, provided_password):
        """Verify a stored password against a provided password."""
        salt = stored_password.split(':')[1]
        stored_hash = stored_password.split(':')[0]
        hash_of_provided = hashlib.sha256(salt.encode() + provided_password.encode()).hexdigest()
        return hash_of_provided == stored_hash

    def authenticate_user(self, username, password):
        """Authenticate a user with username and password."""
        try:
            # First check if the user exists
            self.cursor.execute('''
            SELECT id, username, password_hash, full_name, email, role_id
            FROM users
            WHERE username = ?
            ''', (username,))

            user_data = self.cursor.fetchone()
            if not user_data:
                print(f"User not found: {username}")
                return None

            user_id, user_name, password_hash, full_name, email, role_id = user_data

            # Then get the role information
            if role_id:
                self.cursor.execute('''
                SELECT name, permissions
                FROM roles
                WHERE id = ?
                ''', (role_id,))

                role_data = self.cursor.fetchone()
                if not role_data:
                    print(f"Role not found for role_id: {role_id}")
                    return None

                role_name, permissions = role_data
            else:
                role_name = "anonymous"
                permissions = ""

            # Verify password
            if self._verify_password(password_hash, password):
                # Update last login time
                self.cursor.execute('''
                UPDATE users SET last_login = ? WHERE id = ?
                ''', (datetime.now().isoformat(), user_id))
                self.conn.commit()

                # Set current user
                self.current_user = {
                    'id': user_id,
                    'username': user_name,
                    'full_name': full_name,
                    'email': email,
                    'role_id': role_id,
                    'role_name': role_name,
                    'permissions': permissions.split(',') if permissions else []
                }
                return self.current_user

            print(f"Invalid password for user: {username}")
            return None
        except sqlite3.Error as e:
            print(f"Authentication error: {e}")
            import traceback
            traceback.print_exc()
            return None

    def get_current_user(self):
        """Get the currently authenticated user."""
        return self.current_user

    def logout(self):
        """Log out the current user."""
        self.current_user = None

    def get_all_users(self):
        """Get all users from the database."""
        try:
            self.cursor.execute('''
            SELECT id, username, full_name, email, role_id, created_at, last_login
            FROM users
            ORDER BY username
            ''')
            users = self.cursor.fetchall()
            return [
                {
                    'id': u[0],
                    'username': u[1],
                    'full_name': u[2],
                    'email': u[3],
                    'role_id': u[4],
                    'created_at': u[5],
                    'last_login': u[6]
                } for u in users
            ]
        except sqlite3.Error as e:
            print(f"Error getting users: {e}")
            return []

    def get_user_by_id(self, user_id):
        """Get a user by ID."""
        try:
            self.cursor.execute('''
            SELECT id, username, full_name, email, role_id, created_at, last_login
            FROM users
            WHERE id = ?
            ''', (user_id,))
            user = self.cursor.fetchone()
            if user:
                return {
                    'id': user[0],
                    'username': user[1],
                    'full_name': user[2],
                    'email': user[3],
                    'role_id': user[4],
                    'created_at': user[5],
                    'last_login': user[6]
                }
            return None
        except sqlite3.Error as e:
            print(f"Error getting user: {e}")
            return None

    def create_user(self, username, password, email, full_name, role_id):
        """Create a new user."""
        try:
            password_hash = self._hash_password(password)
            self.cursor.execute('''
            INSERT INTO users (username, password_hash, email, full_name, role_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', (username, password_hash, email, full_name, role_id, datetime.now().isoformat()))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error creating user: {e}")
            return False

    def update_user(self, user_id, full_name, email, role_id, new_password=None):
        """Update an existing user."""
        try:
            if new_password:
                # Update with new password
                password_hash = self._hash_password(new_password)
                self.cursor.execute('''
                UPDATE users
                SET full_name = ?, email = ?, role_id = ?, password_hash = ?
                WHERE id = ?
                ''', (full_name, email, role_id, password_hash, user_id))
            else:
                # Update without changing password
                self.cursor.execute('''
                UPDATE users
                SET full_name = ?, email = ?, role_id = ?
                WHERE id = ?
                ''', (full_name, email, role_id, user_id))

            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error updating user: {e}")
            return False

    def delete_user(self, user_id):
        """Delete a user."""
        try:
            self.cursor.execute('''
            DELETE FROM users
            WHERE id = ?
            ''', (user_id,))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error deleting user: {e}")
            return False

    def create_role(self, name, description, permissions):
        """Create a new role."""
        try:
            self.cursor.execute('''
            INSERT INTO roles (name, description, permissions)
            VALUES (?, ?, ?)
            ''', (name, description, permissions))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error creating role: {e}")
            return False

    def update_role(self, role_id, name, description, permissions):
        """Update an existing role."""
        try:
            self.cursor.execute('''
            UPDATE roles
            SET name = ?, description = ?, permissions = ?
            WHERE id = ?
            ''', (name, description, permissions, role_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error updating role: {e}")
            return False

    def delete_role(self, role_id):
        """Delete a role."""
        try:
            # Check if any users have this role
            self.cursor.execute('SELECT COUNT(*) FROM users WHERE role_id = ?', (role_id,))
            if self.cursor.fetchone()[0] > 0:
                # Set all users with this role to have no role
                self.cursor.execute('UPDATE users SET role_id = NULL WHERE role_id = ?', (role_id,))

            # Delete the role
            self.cursor.execute('DELETE FROM roles WHERE id = ?', (role_id,))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error deleting role: {e}")
            return False

    def get_roles(self):
        """Get all available roles."""
        try:
            self.cursor.execute('''
            SELECT id, name, description, permissions FROM roles
            ORDER BY name
            ''')
            roles = self.cursor.fetchall()
            return [{'id': r[0], 'name': r[1], 'description': r[2], 'permissions': r[3].split(',')} for r in roles]
        except sqlite3.Error as e:
            print(f"Error getting roles: {e}")
            return []

    def has_permission(self, permission):
        """Check if the current user has a specific permission."""
        if not self.current_user or 'permissions' not in self.current_user:
            return False

        return permission in self.current_user['permissions']

    def close(self):
        """Close the database connection."""
        try:
            # First close the cursor if it exists
            if hasattr(self, 'cursor') and self.cursor:
                try:
                    self.cursor.close()
                    print("User management cursor closed")
                except Exception as cursor_e:
                    print(f"Error closing user management cursor: {str(cursor_e)}")

            # Then close the connection
            if hasattr(self, 'conn') and self.conn:
                try:
                    self.conn.close()
                    print("User management connection closed")
                except Exception as conn_e:
                    print(f"Error closing user management connection: {str(conn_e)}")

            # Clear references to prevent further use
            self.cursor = None
            self.conn = None
            print("User management database connection closed successfully")
        except Exception as e:
            print(f"Error in user management close method: {str(e)}")
            import traceback
            traceback.print_exc()
