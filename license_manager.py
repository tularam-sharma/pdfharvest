import uuid
import hashlib
import platform
import json
import os
import datetime
import socket
import requests
import shutil
import zlib
import base64
import sqlite3
from typing import Dict, Any, Optional, Tuple

class LicenseManager:
    """
    License manager for hardware-locked application protection
    """
    def __init__(self, app_id: str = "pdf_extractor_app", license_file: str = "license.key", db_file: str = "user_management.db"):
        self.app_id = app_id
        self.license_file = license_file
        self.db_file = db_file
        self.hardware_id = self._generate_hardware_id()
        self.current_user_id = None  # Will be set when a user logs in

    def _generate_hardware_id(self) -> str:
        """
        Generate a unique hardware ID based on system components
        """
        # Get system information
        system_info = {
            "machine_id": str(uuid.getnode()),  # MAC address as integer
            "platform": platform.platform(),
            "processor": platform.processor(),
            "hostname": socket.gethostname(),
            "username": os.getlogin()
        }

        # Create a stable hardware ID by hashing system information
        hardware_id_str = f"{system_info['machine_id']}:{system_info['platform']}:{system_info['processor']}"
        hardware_id = hashlib.sha256(hardware_id_str.encode()).hexdigest()

        return hardware_id

    def generate_license_request(self) -> Dict[str, Any]:
        """
        Generate a license request that can be sent to a license server
        """
        request_data = {
            "app_id": self.app_id,
            "hardware_id": self.hardware_id,
            "hostname": socket.gethostname(),
            "username": os.getlogin(),
            "timestamp": datetime.datetime.now().isoformat(),
            "platform": platform.platform(),
            "processor": platform.processor()
        }

        return request_data

    def save_license_request(self, filename: str = "license_request.json") -> str:
        """
        Save the license request to a file that can be sent to the vendor
        """
        request_data = self.generate_license_request()

        with open(filename, "w") as f:
            json.dump(request_data, f, indent=2)

        return filename

    def activate_with_key(self, license_key: str) -> Tuple[bool, str]:
        """
        Activate the software with the provided license key
        Returns (success, message)
        """
        # Verify the license key format (should be a valid signature)
        if not self._is_valid_key_format(license_key):
            return False, "Invalid license key format"

        # Decode and verify the license key
        try:
            license_data = self._decode_license_key(license_key)
            if not license_data:
                return False, "Invalid license key"

            # Verify hardware ID if present in the license
            if "hardware_id" in license_data:
                # The license may contain a truncated hardware ID, so we need to check if
                # the license hardware ID is a prefix of the actual hardware ID
                license_hw_id = license_data["hardware_id"]
                if not self.hardware_id.startswith(license_hw_id):
                    return False, "This license key is not valid for this computer"

            # Check license expiration
            if "expiry_date" in license_data:
                try:
                    expiry_date = datetime.datetime.fromisoformat(license_data["expiry_date"])
                    if expiry_date < datetime.datetime.now():
                        return False, f"License expired on {expiry_date.strftime('%Y-%m-%d')}"
                except (ValueError, TypeError):
                    pass  # Invalid date format, will be handled by general validation

            # Add activation data
            license_data["activation_date"] = datetime.datetime.now().isoformat()
            license_data["hardware_id"] = self.hardware_id  # Always store current hardware ID
            license_data["app_id"] = self.app_id
            license_data["key"] = license_key

            # Save the license file
            with open(self.license_file, "w") as f:
                json.dump(license_data, f)

            # Create a success message based on license features
            message = "License activated successfully"

            # Add expiration info if available
            if "expiry_date" in license_data:
                try:
                    expiry_date = datetime.datetime.fromisoformat(license_data["expiry_date"])
                    message += f". Valid until {expiry_date.strftime('%Y-%m-%d')}"
                except (ValueError, TypeError):
                    pass

            # Add edition info if available
            if "edition" in license_data:
                message += f". {license_data['edition']} Edition"

            return True, message

        except Exception as e:
            return False, f"Activation error: {str(e)}"

    def _is_valid_key_format(self, key: str) -> bool:
        """
        Check if the license key has a valid format
        """
        # Basic format validation
        if not key or len(key) < 32 or "-" not in key:
            return False

        # Check if it's in our expected format (blocks of chars separated by hyphens)
        parts = key.split("-")

        # Must have at least 4 parts
        if len(parts) < 4:
            return False

        # First part should be 8 characters (checksum)
        if len(parts[0]) != 8:
            return False

        return True

    def _decode_license_key(self, key: str) -> Dict[str, Any]:
        """
        Decode a license key into its component data
        """
        if not self._is_valid_key_format(key):
            return {}

        try:
            # Remove hyphens
            clean_key = key.replace("-", "")

            # The first 8 characters are a checksum
            checksum = clean_key[:8]
            encoded_data = clean_key[8:]

            # Verify the checksum
            calculated_checksum = hashlib.md5(encoded_data.encode()).hexdigest()[:8]
            if calculated_checksum.lower() != checksum.lower():
                return {}

            # Try to decode using the new method (base64 + zlib)
            try:
                import base64
                import zlib

                # Add padding if needed
                padding_needed = len(encoded_data) % 4
                if padding_needed:
                    encoded_data += "=" * (4 - padding_needed)

                # Convert from base64 to bytes
                try:
                    # First try standard base64
                    compressed_data = base64.b64decode(encoded_data)
                except Exception:
                    # Fall back to URL-safe base64
                    compressed_data = base64.urlsafe_b64decode(encoded_data)

                # Decompress
                json_bytes = zlib.decompress(compressed_data)

                # Decode JSON
                compact_data = json.loads(json_bytes.decode('utf-8'))

                # Convert compact data back to full license data
                license_data = {}

                # Edition
                if "e" in compact_data:
                    license_data["edition"] = compact_data["e"]

                # Expiry date
                if "d" in compact_data:
                    date_str = compact_data["d"]
                    if len(date_str) == 8:  # YYYYMMDD format
                        try:
                            year = int(date_str[:4])
                            month = int(date_str[4:6])
                            day = int(date_str[6:8])
                            license_data["expiry_date"] = datetime.datetime(year, month, day).isoformat()
                        except ValueError:
                            license_data["expiry_date"] = date_str
                    else:
                        license_data["expiry_date"] = date_str

                # File limit
                if "f" in compact_data:
                    license_data["file_limit"] = compact_data["f"]

                # Features
                if "ft" in compact_data:
                    # Convert compact features back to full names
                    feature_map = {
                        "b": "basic_extraction",
                        "e": "export_data",
                        "v": "view_templates",
                        "a": "advanced_extraction",
                        "p": "batch_processing",
                        "i": "api_access",
                        "c": "custom_templates"
                    }
                    license_data["features"] = [feature_map.get(f, f) for f in compact_data["ft"]]

                # Hardware ID
                if "h" in compact_data:
                    license_data["hardware_id"] = compact_data["h"]

                # Creation timestamp
                if "t" in compact_data:
                    creation_time = datetime.datetime.fromtimestamp(compact_data["t"])
                    license_data["creation_date"] = creation_time.isoformat()

                return license_data

            except Exception as e:
                # Fall back to the old method
                pass

            # Try the old hex method as fallback
            try:
                # Convert the hex string to bytes
                data_bytes = bytes.fromhex(encoded_data)
                # Decode as JSON
                decoded_str = data_bytes.decode('utf-8')
                license_data = json.loads(decoded_str)
                return license_data
            except Exception:
                # If that fails, try our fallback method for demo keys
                if encoded_data.startswith("DEMO"):
                    # Demo license format
                    edition = "Demo"
                    # Extract expiry date (YYYYMMDD format)
                    date_str = encoded_data[4:12]
                    try:
                        year = int(date_str[:4])
                        month = int(date_str[4:6])
                        day = int(date_str[6:8])
                        expiry_date = datetime.datetime(year, month, day).isoformat()
                    except ValueError:
                        # Default to 30 days from now if date is invalid
                        expiry_date = (datetime.datetime.now() + datetime.timedelta(days=30)).isoformat()

                    # Extract file limit if present
                    file_limit = 10  # Default
                    if len(encoded_data) > 12:
                        try:
                            file_limit = int(encoded_data[12:16])
                        except ValueError:
                            pass

                    return {
                        "edition": edition,
                        "expiry_date": expiry_date,
                        "file_limit": file_limit,
                        "features": ["basic_extraction"],
                        "is_demo": True
                    }
                return {}

        except Exception:
            return {}

    def verify_license(self) -> Tuple[bool, str]:
        """
        Verify that the current license is valid for this hardware
        Returns (is_valid, message)
        """
        # Check if license file exists
        if not os.path.exists(self.license_file):
            return False, "No license found. Please activate the software."

        try:
            # Read license data
            with open(self.license_file, "r") as f:
                license_data = json.load(f)

            # Verify app ID
            if license_data.get("app_id") != self.app_id:
                return False, "Invalid license: Application ID mismatch."

            # Verify hardware ID
            if "hardware_id" in license_data:
                # The license may contain a truncated hardware ID, so we need to check if
                # the license hardware ID is a prefix of the actual hardware ID
                license_hw_id = license_data["hardware_id"]
                if not self.hardware_id.startswith(license_hw_id):
                    return False, "Invalid license: Hardware ID mismatch. This license is not valid for this computer."

            # Check license expiration
            if "expiry_date" in license_data:
                try:
                    expiry_date = datetime.datetime.fromisoformat(license_data["expiry_date"])
                    if expiry_date < datetime.datetime.now():
                        return False, f"License expired on {expiry_date.strftime('%Y-%m-%d')}. Please renew your license."

                    # If expiration is within 30 days, add a warning
                    days_remaining = (expiry_date - datetime.datetime.now()).days
                    if days_remaining <= 30:
                        return True, f"License valid, but expires in {days_remaining} days. Please renew soon."
                except (ValueError, TypeError):
                    pass

            # Create a success message based on license features
            message = "License verified successfully"

            # Add edition info if available
            if "edition" in license_data:
                message += f". {license_data['edition']} Edition"

            # Add file limit info if available
            if "file_limit" in license_data:
                message += f". Bulk extraction limit: {license_data['file_limit']} files"

            # Add expiration info if available
            if "expiry_date" in license_data:
                try:
                    expiry_date = datetime.datetime.fromisoformat(license_data["expiry_date"])
                    message += f". Valid until {expiry_date.strftime('%Y-%m-%d')}"
                except (ValueError, TypeError):
                    pass

            return True, message

        except Exception as e:
            return False, f"License verification error: {str(e)}"

    def remove_license(self) -> Tuple[bool, str]:
        """
        Remove the current license file
        Returns (success, message)
        """
        if not os.path.exists(self.license_file):
            return False, "No license found to remove."

        try:
            # Create a backup of the license file
            backup_file = f"{self.license_file}.bak"
            if os.path.exists(self.license_file):
                # Read the current license data for the message
                with open(self.license_file, "r") as f:
                    license_data = json.load(f)

                # Create a backup
                shutil.copy2(self.license_file, backup_file)

                # Remove the license file
                os.remove(self.license_file)

                # Get the activation date for the message
                activation_date = license_data.get("activation_date", "unknown date")
                return True, f"License successfully removed. The license was activated on {activation_date}."

        except Exception as e:
            return False, f"Error removing license: {str(e)}"

        return False, "Failed to remove license for unknown reason."

    def get_license_info(self) -> Dict[str, Any]:
        """
        Get information about the current license
        Returns a dictionary with license details or an empty dict if no license
        """
        if not os.path.exists(self.license_file):
            return {}

        try:
            # Read license data
            with open(self.license_file, "r") as f:
                license_data = json.load(f)

            # Add some computed fields
            if "expiry_date" in license_data:
                try:
                    expiry_date = datetime.datetime.fromisoformat(license_data["expiry_date"])
                    license_data["days_remaining"] = (expiry_date - datetime.datetime.now()).days
                    license_data["is_expired"] = expiry_date < datetime.datetime.now()
                except (ValueError, TypeError):
                    license_data["days_remaining"] = 0
                    license_data["is_expired"] = True

            return license_data

        except Exception:
            return {}

    def check_feature_access(self, feature_name: str) -> Tuple[bool, str]:
        """
        Check if the current license allows access to a specific feature
        Returns (has_access, message)
        """
        # First verify the license is valid
        is_valid, message = self.verify_license()
        if not is_valid:
            return False, message

        # Get license info
        license_info = self.get_license_info()

        # Check if the license has a features list
        features = license_info.get("features", [])

        # Some features are always available with a valid license
        basic_features = ["basic_extraction", "view_templates", "export_data"]

        # Check if the feature is available
        if feature_name in basic_features or feature_name in features:
            return True, "Feature access granted"

        # Special handling for edition-based features
        edition = license_info.get("edition", "").lower()

        # Professional edition features
        if edition in ["professional", "enterprise"] and feature_name in ["advanced_extraction", "batch_processing"]:
            return True, "Feature access granted"

        # Enterprise edition features
        if edition == "enterprise" and feature_name in ["api_access", "custom_templates"]:
            return True, "Feature access granted"

        return False, f"Your license does not include access to {feature_name}. Please upgrade your license."

    def check_bulk_limit(self, file_count: int) -> Tuple[bool, str]:
        """
        Check if the current license allows processing the specified number of files
        Returns (is_allowed, message)
        """
        # First verify the license is valid
        is_valid, message = self.verify_license()
        if not is_valid:
            return False, message

        # Get license info
        license_info = self.get_license_info()

        # Check if the license has a file limit
        file_limit = license_info.get("file_limit", 0)

        # If no limit is specified (0 or not present), or the limit is -1 (unlimited)
        if file_limit <= 0 or file_limit == -1:
            return True, "No file limit restriction"

        # Get the current user's file processing count
        files_processed = self.get_files_processed()

        # Check if the requested count exceeds the limit
        if files_processed + file_count > file_limit:
            return False, f"File count ({file_count}) would exceed your license limit of {file_limit} files. You have processed {files_processed} files already. Please upgrade your license or process fewer files."

        return True, f"Within file limit ({files_processed}/{file_limit})"

    def online_activation(self, license_key: str, activation_server: str) -> Tuple[bool, str]:
        """
        Activate the software online by contacting an activation server
        """
        try:
            # Prepare activation data
            activation_data = {
                "license_key": license_key,
                "hardware_id": self.hardware_id,
                "app_id": self.app_id,
                "hostname": socket.gethostname(),
                "platform": platform.platform()
            }

            # Send activation request
            response = requests.post(
                f"{activation_server}/activate",
                json=activation_data,
                timeout=10
            )

            if response.status_code == 200:
                result = response.json()
                if result.get("status") == "success":
                    # Save the license locally
                    self.activate_with_key(license_key)
                    return True, "Activation successful."
                else:
                    return False, f"Activation failed: {result.get('message', 'Unknown error')}"
            else:
                return False, f"Activation server error: {response.status_code}"

        except Exception as e:
            return False, f"Activation error: {str(e)}"

    def set_current_user(self, user_id: int) -> None:
        """
        Set the current user ID for file tracking
        """
        self.current_user_id = user_id

    def get_files_processed(self) -> int:
        """
        Get the number of files processed by the current user
        """
        if not self.current_user_id:
            return 0

        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            cursor.execute(
                "SELECT files_processed FROM file_tracking WHERE user_id = ?",
                (self.current_user_id,)
            )
            result = cursor.fetchone()

            conn.close()

            if result:
                return result[0]
            return 0
        except Exception as e:
            print(f"Error getting files processed: {str(e)}")
            return 0

    def update_files_processed(self, additional_files: int) -> bool:
        """
        Update the number of files processed by the current user
        """
        if not self.current_user_id:
            return False

        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            # Get current count
            cursor.execute(
                "SELECT files_processed FROM file_tracking WHERE user_id = ?",
                (self.current_user_id,)
            )
            result = cursor.fetchone()

            if result:
                current_count = result[0]
                new_count = current_count + additional_files

                # Update the count
                cursor.execute(
                    "UPDATE file_tracking SET files_processed = ?, last_updated = ? WHERE user_id = ?",
                    (new_count, datetime.datetime.now().isoformat(), self.current_user_id)
                )
                conn.commit()
                conn.close()
                return True
            else:
                # Insert a new record if one doesn't exist
                cursor.execute(
                    "INSERT INTO file_tracking (user_id, files_processed, last_updated) VALUES (?, ?, ?)",
                    (self.current_user_id, additional_files, datetime.datetime.now().isoformat())
                )
                conn.commit()
                conn.close()
                return True
        except Exception as e:
            print(f"Error updating files processed: {str(e)}")
            return False

    def reset_files_processed(self, user_id: int = None) -> bool:
        """
        Reset the number of files processed by a user or all users
        """
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            if user_id:
                # Reset for a specific user
                cursor.execute(
                    "UPDATE file_tracking SET files_processed = 0, last_updated = ? WHERE user_id = ?",
                    (datetime.datetime.now().isoformat(), user_id)
                )
            else:
                # Reset for all users
                cursor.execute(
                    "UPDATE file_tracking SET files_processed = 0, last_updated = ?",
                    (datetime.datetime.now().isoformat(),)
                )

            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error resetting files processed: {str(e)}")
            return False

    def get_all_user_file_counts(self) -> list:
        """
        Get file processing counts for all users
        Returns a list of tuples (user_id, username, files_processed, last_updated)
        """
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            cursor.execute("""
                SELECT u.id, u.username, u.full_name, f.files_processed, f.last_updated
                FROM users u
                LEFT JOIN file_tracking f ON u.id = f.user_id
            """)

            results = cursor.fetchall()
            conn.close()

            return results
        except Exception as e:
            print(f"Error getting user file counts: {str(e)}")
            return []


# Singleton instance
_license_manager = None

def get_license_manager(app_id: str = "pdf_extractor_app", user_id: int = 1) -> LicenseManager:
    """
    Get the singleton instance of the license manager
    """
    global _license_manager
    if _license_manager is None:
        _license_manager = LicenseManager(app_id=app_id)
        # Set the default user ID (usually admin)
        _license_manager.set_current_user(user_id)
    return _license_manager
