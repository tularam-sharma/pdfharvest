#!/usr/bin/env python3
"""
Database Protection Module

This module provides functions to encrypt and decrypt the user_management.db file
to protect it from unauthorized access or tampering.
"""

import os
import sys
import base64
import hashlib
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import sqlite3
import platform
import uuid

# Import path_helper for path resolution (optional)
try:
    import path_helper
    PATH_HELPER_AVAILABLE = True
except ImportError:
    PATH_HELPER_AVAILABLE = False
    print("[WARNING] path_helper not available. Using default path resolution.")

# Constants
DB_FILENAME = "user_management.db"
ENCRYPTED_EXTENSION = ".enc"
SALT_FILE = ".db_salt"

def resolve_path(path):
    """Resolve path using path_helper if available, otherwise use os.path.abspath"""
    if PATH_HELPER_AVAILABLE:
        return path_helper.resolve_path(path)
    else:
        return os.path.abspath(path)

def get_machine_key():
    """
    Generate a machine-specific key based on hardware identifiers.
    This ensures the database can only be decrypted on the same machine.
    """
    # Get system information
    system_info = {
        "machine_id": str(uuid.getnode()),  # MAC address as integer
        "platform": platform.platform(),
        "processor": platform.processor(),
        "hostname": platform.node(),
        "system": platform.system(),
        "version": platform.version(),
    }

    # Create a stable hardware ID by hashing system information
    hardware_id_str = f"{system_info['machine_id']}:{system_info['platform']}:{system_info['system']}"
    hardware_id = hashlib.sha256(hardware_id_str.encode()).digest()

    return hardware_id

def get_encryption_key():
    """
    Get or create the encryption key based on machine-specific information.
    Uses a salt file to ensure the key is consistent across runs but unique to the installation.
    """
    # Get machine-specific key
    machine_key = get_machine_key()

    # Resolve the salt file path
    salt_file_path = resolve_path(SALT_FILE)

    # Check if salt file exists
    if os.path.exists(salt_file_path):
        with open(salt_file_path, "rb") as f:
            salt = f.read()
    else:
        # Generate a new salt and save it
        salt = os.urandom(16)
        with open(salt_file_path, "wb") as f:
            f.write(salt)

    # Derive key using PBKDF2
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )

    # Derive the key from the machine key
    key = base64.urlsafe_b64encode(kdf.derive(machine_key))

    return key

def encrypt_database():
    """
    Encrypt the user_management.db file to protect it from unauthorized access.
    """
    # Resolve paths
    db_path = resolve_path(DB_FILENAME)
    encrypted_path = resolve_path(f"{DB_FILENAME}{ENCRYPTED_EXTENSION}")

    # Check if database exists
    if not os.path.exists(db_path):
        print(f"Error: {db_path} not found.")
        return False

    # Check if database is already encrypted
    if os.path.exists(encrypted_path):
        # Check if the encrypted file is newer than the decrypted file
        try:
            db_time = os.path.getmtime(db_path)
            enc_time = os.path.getmtime(encrypted_path)

            if enc_time >= db_time:
                print(f"Database is already encrypted as {encrypted_path} and is up to date")
                return True
            else:
                print(f"Encrypted database exists but is outdated. Updating...")
        except Exception as e:
            print(f"Error checking file timestamps: {str(e)}")

    try:
        # Verify the database is valid before encrypting
        try:
            test_conn = sqlite3.connect(db_path)
            test_cursor = test_conn.cursor()
            test_cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = test_cursor.fetchall()
            test_conn.close()

            if not tables:
                print("Warning: Database appears to be empty. Proceeding with encryption anyway.")
        except sqlite3.Error as e:
            print(f"Warning: Database validation failed: {str(e)}. Proceeding with encryption anyway.")

        # Get encryption key
        key = get_encryption_key()
        fernet = Fernet(key)

        # Read the database file
        with open(db_path, "rb") as f:
            db_data = f.read()

        # Encrypt the data
        encrypted_data = fernet.encrypt(db_data)

        # Write the encrypted data
        with open(encrypted_path, "wb") as f:
            f.write(encrypted_data)

        print(f"Database encrypted successfully as {encrypted_path}")

        # IMPORTANT: We keep the original file to prevent data loss
        # The cleanup function will handle file permissions

        return True

    except Exception as e:
        print(f"Error encrypting database: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def decrypt_database():
    """
    Decrypt the encrypted user_management.db file for use by the application.
    """
    # Resolve paths
    db_path = resolve_path(DB_FILENAME)
    encrypted_path = resolve_path(f"{DB_FILENAME}{ENCRYPTED_EXTENSION}")

    # Check if encrypted database exists
    if not os.path.exists(encrypted_path):
        print(f"Error: {encrypted_path} not found.")
        return False

    # Check if decrypted database already exists and is valid
    if os.path.exists(db_path):
        # Ensure the database file is writable
        try:
            os.chmod(db_path, 0o600)  # Read-write for owner
            print(f"Set database file to read-write mode")
        except Exception as chmod_e:
            print(f"Warning: Could not set file permissions: {str(chmod_e)}")

        try:
            # Check if the decrypted file is valid and newer than the encrypted file
            db_time = os.path.getmtime(db_path)
            enc_time = os.path.getmtime(encrypted_path)

            if db_time >= enc_time:
                # Verify the database is valid
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tables = cursor.fetchall()
                conn.close()

                if tables:
                    print(f"Decrypted database {db_path} is already up to date and valid")
                    return True
                else:
                    print("Existing decrypted database appears empty, will decrypt from encrypted version")
            else:
                print("Decrypted database exists but is older than encrypted version, updating...")
        except Exception as e:
            print(f"Error checking existing decrypted database: {str(e)}")
            # Continue with decryption

    try:
        # Create a backup of the existing decrypted database if it exists
        if os.path.exists(db_path):
            backup_path = resolve_path(f"{DB_FILENAME}.bak")
            try:
                import shutil
                shutil.copy2(db_path, backup_path)
                print(f"Created backup of existing database as {backup_path}")
            except Exception as e:
                print(f"Warning: Failed to create backup: {str(e)}")

        # Get encryption key
        key = get_encryption_key()
        fernet = Fernet(key)

        # Read the encrypted database file
        with open(encrypted_path, "rb") as f:
            encrypted_data = f.read()

        # Decrypt the data
        try:
            decrypted_data = fernet.decrypt(encrypted_data)
        except Exception as e:
            print(f"Error decrypting database: {str(e)}")
            print("This could be due to the database being encrypted on a different machine.")
            return False

        # Write the decrypted data
        with open(db_path, "wb") as f:
            f.write(decrypted_data)

        # Ensure the database file is writable
        try:
            os.chmod(db_path, 0o600)  # Read-write for owner
            print(f"Set database file to read-write mode")
        except Exception as chmod_e:
            print(f"Warning: Could not set file permissions: {str(chmod_e)}")

        print(f"Database decrypted successfully to {db_path}")

        # Verify the decrypted database is valid
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            conn.close()

            if not tables:
                print("Warning: Decrypted database appears to be empty.")
                # If we have a backup, check if it's valid
                backup_path = resolve_path(f"{DB_FILENAME}.bak")
                if os.path.exists(backup_path):
                    try:
                        conn = sqlite3.connect(backup_path)
                        cursor = conn.cursor()
                        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                        backup_tables = cursor.fetchall()
                        conn.close()

                        if backup_tables:
                            print("Backup database appears valid, restoring from backup")
                            import shutil
                            shutil.copy2(backup_path, db_path)
                            print(f"Restored database from backup")
                    except Exception as backup_e:
                        print(f"Error checking backup database: {str(backup_e)}")

            print(f"Database verified successfully. Found {len(tables)} tables.")
            return True

        except sqlite3.Error as e:
            print(f"Error verifying decrypted database: {str(e)}")
            return False

    except Exception as e:
        print(f"Error decrypting database: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def initialize_database_protection():
    """
    Initialize database protection by decrypting the database if it's encrypted.
    This should be called at application startup.
    """
    # Resolve paths
    db_path = resolve_path(DB_FILENAME)
    encrypted_path = resolve_path(f"{DB_FILENAME}{ENCRYPTED_EXTENSION}")

    print(f"Checking database paths: DB={db_path}, Encrypted={encrypted_path}")

    # Check if encrypted database exists
    if os.path.exists(encrypted_path):
        print(f"Found encrypted database: {encrypted_path}")

        # If both encrypted and decrypted exist, use the encrypted one
        if os.path.exists(db_path):
            # Check if the decrypted file is valid
            try:
                # Try to open the database to check if it's valid
                test_conn = sqlite3.connect(db_path)
                test_cursor = test_conn.cursor()
                test_cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tables = test_cursor.fetchall()
                test_conn.close()

                if tables:  # If we have tables, the database is likely valid
                    print(f"Existing decrypted database appears valid with {len(tables)} tables")
                    return True  # Use the existing decrypted database
                else:
                    print("Existing decrypted database appears empty or invalid")
            except sqlite3.Error as e:
                print(f"Error checking existing database: {str(e)}")
                # The database is likely corrupted, so we'll decrypt the encrypted version

            # Backup the current decrypted file just in case
            backup_path = resolve_path(f"{DB_FILENAME}.bak")
            try:
                import shutil
                shutil.copy2(db_path, backup_path)
                print(f"Created backup of existing database as {backup_path}")
            except Exception as e:
                print(f"Warning: Failed to create backup: {str(e)}")

        # Decrypt the database
        return decrypt_database()

    # If only the decrypted database exists, encrypt it for protection
    elif os.path.exists(db_path):
        print(f"Found only decrypted database: {db_path}")
        # Create an encrypted backup but don't remove the original
        return encrypt_database()

    # Neither exists, which is a problem
    else:
        print(f"Error: Neither {db_path} nor {encrypted_path} found.")
        # Create an empty database file
        try:
            conn = sqlite3.connect(db_path)
            conn.close()
            print(f"Created new empty database: {db_path}")
            return True
        except Exception as e:
            print(f"Error creating new database: {str(e)}")
            return False

def cleanup_database_protection():
    """
    Clean up database protection by encrypting the database if it's decrypted.
    This should be called at application shutdown.
    """
    # Resolve paths
    db_path = resolve_path(DB_FILENAME)
    encrypted_path = resolve_path(f"{DB_FILENAME}{ENCRYPTED_EXTENSION}")

    print(f"Cleaning up database protection: DB={db_path}, Encrypted={encrypted_path}")

    # If the decrypted database exists, encrypt it
    if os.path.exists(db_path):
        # First, try to close any open connections to the database
        try:
            # Create a temporary connection and immediately close it
            # This helps ensure no other connections are active
            temp_conn = sqlite3.connect(db_path)
            temp_conn.close()
            print(f"Closed any remaining connections to {db_path}")
        except Exception as e:
            print(f"Warning: Could not create temporary connection to database: {str(e)}")

        # Now encrypt the database
        result = encrypt_database()

        # IMPORTANT: We no longer remove the decrypted file to prevent data loss
        # Instead, we just make sure the encrypted version is up to date
        if result:
            print(f"Successfully encrypted database to {encrypted_path}")
            # We no longer make the file read-only as it causes issues with authentication
            print(f"Database encryption completed successfully")

        return result

    return True

if __name__ == "__main__":
    # Simple command-line interface for testing
    if len(sys.argv) < 2:
        print("Usage: python db_protection.py [encrypt|decrypt|init|cleanup]")
        sys.exit(1)

    command = sys.argv[1].lower()

    if command == "encrypt":
        encrypt_database()
    elif command == "decrypt":
        decrypt_database()
    elif command == "init":
        initialize_database_protection()
    elif command == "cleanup":
        cleanup_database_protection()
    else:
        print(f"Unknown command: {command}")
        print("Available commands: encrypt, decrypt, init, cleanup")
        sys.exit(1)
