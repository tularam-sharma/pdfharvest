import sqlite3
import json
import os
import shutil
import threading
import time
from pathlib import Path
import datetime
import signal
from contextlib import contextmanager

# Import path_helper for path resolution (optional)
try:
    import path_helper
    PATH_HELPER_AVAILABLE = True
except ImportError:
    PATH_HELPER_AVAILABLE = False
    print("[WARNING] path_helper not available. Using default path resolution.")


# Timeout exception class
class DatabaseOperationTimeout(Exception):
    """Exception raised when a database operation times out"""
    pass


@contextmanager
def timeout(seconds, operation_name="Database operation"):
    """Context manager for timing out operations"""
    def timeout_handler(signum, frame):
        raise DatabaseOperationTimeout(f"{operation_name} timed out after {seconds} seconds")

    # Set the timeout handler
    original_handler = signal.signal(signal.SIGALRM, timeout_handler)
    signal.alarm(seconds)

    try:
        yield
    finally:
        # Reset the alarm and restore the original handler
        signal.alarm(0)
        signal.signal(signal.SIGALRM, original_handler)


def check_disk_space(path, required_space_mb=100):
    """Check if there is enough disk space available

    Args:
        path (str): Path to check disk space for
        required_space_mb (int): Required space in MB

    Returns:
        tuple: (bool, str) - Whether there is enough space and a message
    """
    try:
        # Get the directory containing the file
        directory = os.path.dirname(os.path.abspath(path))
        if not directory:
            directory = '.'

        # Get disk usage statistics
        disk_usage = shutil.disk_usage(directory)
        free_space_mb = disk_usage.free / (1024 * 1024)  # Convert to MB

        # Check if there's enough free space
        if free_space_mb < required_space_mb:
            return False, f"Not enough disk space. Required: {required_space_mb} MB, Available: {free_space_mb:.2f} MB"

        return True, f"Sufficient disk space available: {free_space_mb:.2f} MB"
    except Exception as e:
        return False, f"Error checking disk space: {str(e)}"

class InvoiceDatabase:
    def __init__(self, db_path="invoice_templates.db"):
        """Initialize the database connection"""
        if PATH_HELPER_AVAILABLE:
            self.db_path = path_helper.resolve_path(db_path)
        else:
            self.db_path = os.path.abspath(db_path)
        self.conn = None
        self.cursor = None
        self.connect()

    def connect(self, max_retries=3, retry_delay=1.0, connection_timeout=30.0):
        """Establish database connection with retry logic

        Args:
            max_retries (int): Maximum number of connection attempts
            retry_delay (float): Delay between retries in seconds
            connection_timeout (float): SQLite connection timeout in seconds
        """
        last_error = None

        for attempt in range(max_retries):
            try:
                if self.conn is not None:
                    self.close()

                print(f"Connecting to database: {self.db_path} (attempt {attempt+1}/{max_retries})")

                # Check if database file exists
                db_dir = os.path.dirname(self.db_path)
                if db_dir and not os.path.exists(db_dir):
                    os.makedirs(db_dir, exist_ok=True)
                    print(f"Created database directory: {db_dir}")

                # Check disk space before connecting
                has_space, space_message = check_disk_space(self.db_path)
                if not has_space:
                    print(f"WARNING: {space_message}")
                    print("Continuing with connection attempt, but operations may fail.")

                # Connect with timeout
                self.conn = sqlite3.connect(
                    self.db_path,
                    timeout=connection_timeout,
                    isolation_level='IMMEDIATE'  # This provides better concurrency control
                )

                # Set pragmas for better performance and reliability
                self.conn.execute(f"PRAGMA busy_timeout = {int(connection_timeout * 1000)}")  # Convert to milliseconds
                self.conn.execute("PRAGMA journal_mode=WAL")
                self.conn.execute("PRAGMA synchronous=NORMAL")
                self.conn.execute("PRAGMA temp_store=MEMORY")  # Store temp tables in memory for better performance

                # Create cursor and tables
                self.cursor = self.conn.cursor()
                self.create_tables()

                print(f"Successfully connected to database: {self.db_path}")
                return

            except sqlite3.OperationalError as e:
                last_error = e
                error_msg = str(e).lower()

                # Handle specific error cases
                if "database is locked" in error_msg:
                    print(f"Database is locked. Retrying in {retry_delay} seconds...")
                elif "disk i/o error" in error_msg:
                    print(f"Disk I/O error. Check disk health. Retrying in {retry_delay} seconds...")
                elif "unable to open database file" in error_msg:
                    print(f"Unable to open database file. Check permissions. Retrying in {retry_delay} seconds...")
                else:
                    print(f"Database operational error: {str(e)}. Retrying in {retry_delay} seconds...")

                # Wait before retrying
                if attempt < max_retries - 1:  # Don't sleep on the last attempt
                    time.sleep(retry_delay)

            except Exception as e:
                last_error = e
                print(f"Unexpected error connecting to database: {str(e)}")
                if attempt < max_retries - 1:  # Don't sleep on the last attempt
                    time.sleep(retry_delay)

        # If we get here, all retries failed
        print(f"Failed to connect to database after {max_retries} attempts")
        if last_error:
            print(f"Last error: {str(last_error)}")
            import traceback
            traceback.print_exc()
        raise last_error or Exception("Failed to connect to database")

    def create_tables(self):
        """Create necessary tables if they don't exist"""
        try:
            # Create templates table with simplified 3-column schema
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    description TEXT,
                    template_type TEXT DEFAULT 'single',
                    config TEXT,
                    creation_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                    last_modified DATETIME DEFAULT CURRENT_TIMESTAMP,
                    page_count INTEGER DEFAULT 1,
                    json_template TEXT,
                    regions TEXT,
                    column_lines TEXT,
                    extraction_method TEXT DEFAULT 'pypdf_table_extraction'
                )
            """)

            # Add extraction_method column if it doesn't exist
            try:
                self.cursor.execute("ALTER TABLE templates ADD COLUMN extraction_method TEXT DEFAULT 'pypdf_table_extraction'")
                print("Added extraction_method column to templates table")
            except Exception as e:
                if "duplicate column name" in str(e).lower():
                    print("extraction_method column already exists")
                else:
                    print(f"Error adding extraction_method column: {e}")

            self.conn.commit()

        except Exception as e:
            print(f"Error creating tables: {str(e)}")
            raise

    def save_template(self, name, description, config, template_type="single",
                     template_id=None, page_count=1, json_template=None,
                     drawing_regions=None, drawing_column_lines=None, extraction_regions=None,
                     extraction_column_lines=None, drawing_page_regions=None,
                     drawing_page_column_lines=None, extraction_page_regions=None,
                     extraction_page_column_lines=None, extraction_method="pypdf_table_extraction"):
        """
        Save a template to the database with dual coordinate format only.
        If template_id is provided, updates an existing template, otherwise creates a new one.
        """
        try:
            # Import dual coordinate storage system
            from dual_coordinate_storage import DualCoordinateStorage

            # Handle dual coordinate storage
            storage = DualCoordinateStorage()

            # Serialize dual coordinate data
            drawing_regions_json = None
            drawing_column_lines_json = None
            extraction_regions_json = None
            extraction_column_lines_json = None
            drawing_page_regions_json = None
            drawing_page_column_lines_json = None
            extraction_page_regions_json = None
            extraction_page_column_lines_json = None

            if drawing_regions:
                drawing_regions_json = storage.serialize_regions(drawing_regions)
            if drawing_column_lines:
                drawing_column_lines_json = storage.serialize_column_lines(drawing_column_lines)
            if extraction_regions:
                extraction_regions_json = storage.serialize_regions(extraction_regions)
            if extraction_column_lines:
                extraction_column_lines_json = storage.serialize_column_lines(extraction_column_lines)
            if drawing_page_regions:
                drawing_page_regions_json = json.dumps(drawing_page_regions)
            if drawing_page_column_lines:
                drawing_page_column_lines_json = json.dumps(drawing_page_column_lines)
            if extraction_page_regions:
                extraction_page_regions_json = json.dumps(extraction_page_regions)
            if extraction_page_column_lines:
                extraction_page_column_lines_json = json.dumps(extraction_page_column_lines)

            # Serialize config
            config_json = json.dumps(config)

            # Handle JSON template with detailed logging
            print(f"\n[DEBUG] Saving JSON template: {json_template is not None}")
            if json_template:
                print(f"[DEBUG] JSON template type: {type(json_template)}")
                if isinstance(json_template, dict):
                    print(f"[DEBUG] JSON template keys: {list(json_template.keys())}")
                    # Print a sample of the JSON template
                    formatted_json = json.dumps(json_template, indent=2)
                    print(f"[DEBUG] JSON template preview (first 200 chars): {formatted_json[:200]}...")
                try:
                    json_template_str = json.dumps(json_template)
                    print(f"[DEBUG] JSON template serialized successfully (length: {len(json_template_str)})")
                    print(f"[DEBUG] JSON template serialized (first 100 chars): {json_template_str[:100]}...")
                except Exception as json_e:
                    print(f"[ERROR] Failed to serialize JSON template: {str(json_e)}")
                    import traceback
                    traceback.print_exc()
                    json_template_str = None
            else:
                print(f"[DEBUG] JSON template is None or empty")
                json_template_str = None

            # Get current timestamp
            timestamp = datetime.datetime.now().isoformat()

            if template_id:
                # Update existing template
                self.cursor.execute("""
                    UPDATE templates
                    SET name = ?, description = ?, config = ?, template_type = ?, last_modified = ?,
                        page_count = ?, json_template = ?,
                        drawing_regions = ?, drawing_column_lines = ?, extraction_regions = ?,
                        extraction_column_lines = ?, drawing_page_regions = ?, drawing_page_column_lines = ?,
                        extraction_page_regions = ?, extraction_page_column_lines = ?, extraction_method = ?
                    WHERE id = ?
                """, (name, description, config_json, template_type, timestamp,
                      page_count, json_template_str,
                      drawing_regions_json, drawing_column_lines_json, extraction_regions_json,
                      extraction_column_lines_json, drawing_page_regions_json, drawing_page_column_lines_json,
                      extraction_page_regions_json, extraction_page_column_lines_json, extraction_method, template_id))

                print(f"Updated template '{name}' (ID: {template_id})")

            else:
                # Create new template
                self.cursor.execute("""
                    INSERT INTO templates
                    (name, description, config, template_type,
                     creation_date, last_modified, page_count, json_template,
                     drawing_regions, drawing_column_lines, extraction_regions, extraction_column_lines,
                     drawing_page_regions, drawing_page_column_lines, extraction_page_regions, extraction_page_column_lines, extraction_method)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (name, description, config_json, template_type, timestamp, timestamp,
                      page_count, json_template_str,
                      drawing_regions_json, drawing_column_lines_json, extraction_regions_json, extraction_column_lines_json,
                      drawing_page_regions_json, drawing_page_column_lines_json, extraction_page_regions_json, extraction_page_column_lines_json, extraction_method))

                template_id = self.cursor.lastrowid
                print(f"Created new template '{name}' (ID: {template_id})")

            self.conn.commit()
            return template_id

        except Exception as e:
            print(f"Error saving template: {str(e)}")
            import traceback
            traceback.print_exc()
            self.conn.rollback()
            return None

    def get_template(self, template_id=None, template_name=None):
        """
        Get a template by ID or name.
        Returns the template as a dictionary, or None if not found.
        """
        try:
            if template_id:
                self.cursor.execute("""
                    SELECT id, name, description, config, template_type, creation_date, last_modified,
                           page_count, json_template, drawing_regions, drawing_column_lines, extraction_regions,
                           extraction_column_lines, drawing_page_regions, drawing_page_column_lines,
                           extraction_page_regions, extraction_page_column_lines, extraction_method
                    FROM templates
                    WHERE id = ?
                """, (template_id,))
            elif template_name:
                self.cursor.execute("""
                    SELECT id, name, description, config, template_type, creation_date, last_modified,
                           page_count, json_template, drawing_regions, drawing_column_lines, extraction_regions,
                           extraction_column_lines, drawing_page_regions, drawing_page_column_lines,
                           extraction_page_regions, extraction_page_column_lines, extraction_method
                    FROM templates
                    WHERE name = ?
                """, (template_name,))
            else:
                print("Error: Either template_id or template_name must be provided")
                return None

            template = self.cursor.fetchone()
            if not template:
                return None

            # Create template dictionary with dual coordinate support only
            template_dict = {
                'id': template[0],
                'name': template[1],
                'description': template[2],
                'config': json.loads(template[3]),
                'template_type': template[4],
                'creation_date': template[5],
                'last_modified': template[6],
                'page_count': template[7],
                'extraction_method': template[17] if len(template) > 17 and template[17] else 'pypdf_table_extraction'
            }

            # Handle JSON template
            if template[8]:  # json_template exists in database
                print(f"\n[DEBUG] JSON template found in database: {template[8] is not None}")
                print(f"[DEBUG] JSON template raw data length: {len(template[8])}")
                print(f"[DEBUG] JSON template raw data (first 100 chars): {template[8][:100]}...")
                try:
                    json_template = json.loads(template[8])
                    print(f"[DEBUG] JSON template loaded successfully")
                    print(f"[DEBUG] JSON template type: {type(json_template)}")
                    if isinstance(json_template, dict):
                        print(f"[DEBUG] JSON template keys: {list(json_template.keys())}")
                        # Print a sample of the JSON template
                        formatted_json = json.dumps(json_template, indent=2)
                        print(f"[DEBUG] JSON template preview (first 200 chars): {formatted_json[:200]}...")
                    template_dict['json_template'] = json_template
                except Exception as json_e:
                    print(f"[ERROR] Failed to parse JSON template: {str(json_e)}")
                    print(f"[DEBUG] Raw JSON template data: {template[8][:100]}...")
                    import traceback
                    traceback.print_exc()
                    # Create default invoice2data template on error
                    print(f"[DEBUG] Creating default invoice2data template due to parsing error")
                    template_dict['json_template'] = {
                        "issuer": "company_name",
                        "fields": {
                            "invoice_number": "1",
                            "date": "1",
                            "amount": "1"
                        },
                        "keywords": [],
                        "options": {
                            "currency": "EUR",
                            "languages": ["en"],
                            "decimal_separator": ".",
                            "replace": []
                        }
                    }
            else:
                # No JSON template in database, create default invoice2data template
                print(f"\n[DEBUG] No JSON template found in database, creating default invoice2data template")
                template_dict['json_template'] = {
                    "issuer": "company_name",
                    "fields": {
                        "invoice_number": "1",
                        "date": "1",
                        "amount": "1"
                    },
                    "keywords": [],
                    "options": {
                        "currency": "EUR",
                        "languages": ["en"],
                        "decimal_separator": ".",
                        "replace": []
                    }
                }

            # Handle dual coordinate data
            from dual_coordinate_storage import DualCoordinateStorage
            storage = DualCoordinateStorage()

            # Add dual coordinate fields - they are always present in the new schema
            if template[9]:  # drawing_regions
                template_dict['drawing_regions'] = storage.deserialize_regions(template[9])
            if template[10]:  # drawing_column_lines
                template_dict['drawing_column_lines'] = storage.deserialize_column_lines(template[10])
            if template[11]:  # extraction_regions
                template_dict['extraction_regions'] = storage.deserialize_regions(template[11])
            if template[12]:  # extraction_column_lines
                template_dict['extraction_column_lines'] = storage.deserialize_column_lines(template[12])
            if template[13]:  # drawing_page_regions
                template_dict['drawing_page_regions'] = json.loads(template[13])
            if template[14]:  # drawing_page_column_lines
                template_dict['drawing_page_column_lines'] = json.loads(template[14])
            if template[15]:  # extraction_page_regions
                template_dict['extraction_page_regions'] = json.loads(template[15])
            if template[16]:  # extraction_page_column_lines
                template_dict['extraction_page_column_lines'] = json.loads(template[16])

            return template_dict

        except Exception as e:
            print(f"Error getting template: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def delete_template(self, template_id):
        """Delete a template by ID"""
        try:
            self.cursor.execute("DELETE FROM templates WHERE id = ?", (template_id,))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"Error deleting template: {str(e)}")
            self.conn.rollback()
            return False

    def get_all_templates(self):
        """Get all templates from the database"""
        try:
            self.cursor.execute("""
                SELECT id, name, description, template_type, creation_date, last_modified, page_count, json_template, extraction_method
                FROM templates
                ORDER BY creation_date DESC
            """)

            templates = []
            for row in self.cursor.fetchall():
                template_data = {
                    'id': row[0],
                    'name': row[1],
                    'description': row[2],
                    'template_type': row[3],
                    'creation_date': row[4],
                    'last_modified': row[5],
                    'page_count': row[6],
                    'extraction_method': row[8] if len(row) > 8 and row[8] else 'pypdf_table_extraction'
                }

                # Add JSON template if available
                if row[7]:
                    try:
                        print(f"\n[DEBUG] JSON template found in database for template {row[1]}")
                        json_template = json.loads(row[7])
                        print(f"[DEBUG] JSON template loaded successfully")
                        template_data['json_template'] = json_template
                    except Exception as json_e:
                        print(f"[ERROR] Failed to parse JSON template for template {row[1]}: {str(json_e)}")
                        # Create default invoice2data template on error
                        print(f"[DEBUG] Creating default invoice2data template due to parsing error")
                        template_data['json_template'] = {
                            "issuer": "company_name",
                            "fields": {
                                "invoice_number": "1",
                                "date": "1",
                                "amount": "1"
                            },
                            "keywords": [],
                            "options": {
                                "currency": "EUR",
                                "languages": ["en"],
                                "decimal_separator": ".",
                                "replace": []
                            }
                        }
                else:
                    # No JSON template in database, create default invoice2data template
                    print(f"\n[DEBUG] No JSON template found in database for template {row[1]}, creating default invoice2data template")
                    template_data['json_template'] = {
                        "issuer": "company_name",
                        "fields": {
                            "invoice_number": "1",
                            "date": "1",
                            "amount": "1"
                        },
                        "keywords": [],
                        "options": {
                            "currency": "EUR",
                            "languages": ["en"],
                            "decimal_separator": ".",
                            "replace": []
                        }
                    }

                templates.append(template_data)

            return templates

        except Exception as e:
            print(f"Error getting all templates: {str(e)}")
            import traceback
            traceback.print_exc()
            return []

    def close(self):
        """Close the database connection"""
        try:
            if self.cursor:
                try:
                    self.cursor.close()
                except Exception:
                    pass
                self.cursor = None
            if self.conn:
                try:
                    self.conn.commit()  # Ensure all changes are committed
                    self.conn.close()
                except Exception:
                    pass
                self.conn = None
        except Exception as e:
            print(f"Error closing database: {str(e)}")

    def __del__(self):
        """Destructor to ensure connection is closed"""
        self.close()

    def __enter__(self):
        """Context manager entry"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close()

    def execute_with_timeout(self, query, params=None, timeout_seconds=30, operation_name="Database query"):
        """Execute a database query with timeout handling

        Args:
            query (str): SQL query to execute
            params (tuple, optional): Parameters for the query
            timeout_seconds (int): Timeout in seconds
            operation_name (str): Name of the operation for error messages

        Returns:
            cursor: Database cursor after execution

        Raises:
            DatabaseOperationTimeout: If the operation times out
            Exception: For other database errors
        """
        # Windows doesn't support SIGALRM, so we'll use threading instead
        if os.name == 'nt':  # Windows
            result = [None]
            error = [None]
            completed = threading.Event()

            def execute_worker():
                try:
                    if params:
                        result[0] = self.cursor.execute(query, params)
                    else:
                        result[0] = self.cursor.execute(query)
                    completed.set()
                except Exception as e:
                    error[0] = e
                    completed.set()

            # Start the worker thread
            worker_thread = threading.Thread(target=execute_worker)
            worker_thread.daemon = True
            worker_thread.start()

            # Wait for the thread to complete with timeout
            if not completed.wait(timeout_seconds):
                raise DatabaseOperationTimeout(f"{operation_name} timed out after {timeout_seconds} seconds")

            # Check if there was an error
            if error[0] is not None:
                raise error[0]

            return result[0]
        else:  # Unix-like systems
            try:
                with timeout(timeout_seconds, operation_name):
                    if params:
                        return self.cursor.execute(query, params)
                    else:
                        return self.cursor.execute(query)
            except DatabaseOperationTimeout:
                print(f"{operation_name} timed out after {timeout_seconds} seconds")
                raise

    def optimize_database(self, timeout_seconds=300, required_space_factor=2.5):
        """Optimize the database by running VACUUM and ANALYZE

        Args:
            timeout_seconds (int): Maximum time in seconds to allow for optimization
            required_space_factor (float): Factor to multiply database size by to determine required free space

        Returns:
            bool: True if optimization succeeded, False otherwise
        """
        try:
            print("Optimizing database...")

            # Check database size before optimization
            db_size_mb = 0
            if os.path.exists(self.db_path):
                db_size_bytes = os.path.getsize(self.db_path)
                db_size_mb = db_size_bytes / (1024 * 1024)  # Convert to MB
                print(f"Database size before optimization: {db_size_mb:.2f} MB")

            # Calculate required free space (database size * factor)
            required_space_mb = max(100, int(db_size_mb * required_space_factor))

            # Check if there's enough disk space for VACUUM operation
            has_space, space_message = check_disk_space(self.db_path, required_space_mb)
            print(space_message)

            if not has_space:
                print("WARNING: Not enough disk space for safe database optimization.")
                print(f"VACUUM operation requires approximately {required_space_mb} MB of free space.")
                print("Skipping VACUUM operation, but will run other optimizations.")

                # Run PRAGMA optimize (doesn't require as much space)
                self.cursor.execute("PRAGMA optimize")

                # Run ANALYZE to update statistics (doesn't require as much space)
                self.cursor.execute("ANALYZE")

                print("Partial database optimization completed (VACUUM skipped due to disk space constraints)")
                return False

            # Run PRAGMA optimize
            print("Running PRAGMA optimize...")
            self.cursor.execute("PRAGMA optimize")

            # Run ANALYZE to update statistics
            print("Running ANALYZE...")
            self.cursor.execute("ANALYZE")

            # Run VACUUM with timeout
            print(f"Running VACUUM (this may take up to {timeout_seconds} seconds)...")

            # Create a thread to run VACUUM
            vacuum_completed = threading.Event()
            vacuum_error = [None]  # List to store any error that occurs

            def vacuum_worker():
                try:
                    self.conn.execute("VACUUM")
                    vacuum_completed.set()
                except Exception as e:
                    vacuum_error[0] = e
                    vacuum_completed.set()

            # Start the worker thread
            worker_thread = threading.Thread(target=vacuum_worker)
            worker_thread.daemon = True
            worker_thread.start()

            # Wait for the thread to complete with timeout
            vacuum_completed.wait(timeout_seconds)

            # Check if the thread is still running after timeout
            if not vacuum_completed.is_set():
                print(f"VACUUM operation timed out after {timeout_seconds} seconds")
                print("The database may be locked or the operation is taking too long.")
                print("Other optimizations have been applied.")
                return False

            # Check if there was an error
            if vacuum_error[0] is not None:
                print(f"Error during VACUUM operation: {str(vacuum_error[0])}")
                return False

            # Check database size after optimization
            if os.path.exists(self.db_path):
                db_size_after_bytes = os.path.getsize(self.db_path)
                db_size_after_mb = db_size_after_bytes / (1024 * 1024)  # Convert to MB
                print(f"Database size after optimization: {db_size_after_mb:.2f} MB")

                if db_size_mb > 0:  # Avoid division by zero
                    space_saved_mb = db_size_mb - db_size_after_mb
                    space_saved_percent = (space_saved_mb / db_size_mb) * 100 if db_size_mb > 0 else 0
                    print(f"Space saved: {space_saved_mb:.2f} MB ({space_saved_percent:.2f}%)")

            print("Database optimization completed successfully")
            return True
        except Exception as e:
            print(f"Error optimizing database: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def check_integrity(self, repair=True):
        """
        Check the integrity of the database and optionally repair issues.

        Args:
            repair (bool): Whether to attempt repairs if issues are found

        Returns:
            tuple: (bool, str) - Success status and message with details
        """
        try:
            print("Checking database integrity...")

            # Run quick integrity check
            self.cursor.execute("PRAGMA quick_check")
            result = self.cursor.fetchone()
            integrity_status = result[0] if result else "unknown"

            if integrity_status == "ok":
                print("Quick integrity check passed")

                # Run more thorough integrity check
                self.cursor.execute("PRAGMA integrity_check")
                result = self.cursor.fetchall()
                integrity_details = [row[0] for row in result]

                if len(integrity_details) == 1 and integrity_details[0] == "ok":
                    print("Full integrity check passed")
                    return True, "Database integrity verified: no issues found"
                else:
                    error_msg = f"Integrity issues found: {', '.join(integrity_details[:5])}"
                    if len(integrity_details) > 5:
                        error_msg += f" and {len(integrity_details) - 5} more issues"
                    print(error_msg)

                    if repair:
                        return self._repair_database(integrity_details)
                    else:
                        return False, error_msg
            else:
                error_msg = f"Quick integrity check failed: {integrity_status}"
                print(error_msg)

                if repair:
                    return self._repair_database([integrity_status])
                else:
                    return False, error_msg
        except Exception as e:
            error_msg = f"Error checking database integrity: {str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()

            if repair:
                return self._repair_database([str(e)])
            else:
                return False, error_msg

    def _repair_database(self, issues):
        """
        Attempt to repair database integrity issues.

        Args:
            issues (list): List of integrity issues found

        Returns:
            tuple: (bool, str) - Success status and message with details
        """
        print(f"Attempting to repair database with {len(issues)} issues...")

        try:
            # 1. Export all templates to a temporary structure
            templates = []
            exported_count = 0

            try:
                self.cursor.execute("""
                    SELECT id, name, description, regions, column_lines, config,
                           template_type, creation_date, last_modified, page_count,
                           page_regions, page_column_lines, validation_rules, json_template
                    FROM templates
                """)

                for row in self.cursor.fetchall():
                    try:
                        template = {
                            'id': row[0],
                            'name': row[1],
                            'description': row[2],
                            'regions': json.loads(row[3]) if row[3] else {},
                            'column_lines': json.loads(row[4]) if row[4] else {},
                            'config': json.loads(row[5]) if row[5] else {},
                            'template_type': row[6],
                            'creation_date': row[7],
                            'last_modified': row[8],
                            'page_count': row[9],
                            'page_regions': json.loads(row[10]) if row[10] else None,
                            'page_column_lines': json.loads(row[11]) if row[11] else None,
                            'validation_rules': json.loads(row[12]) if row[12] else None,
                            'json_template': json.loads(row[13]) if row[13] else None
                        }
                        templates.append(template)
                        exported_count += 1
                    except json.JSONDecodeError:
                        print(f"Warning: Could not decode JSON for template ID {row[0]}")

                print(f"Exported {exported_count} templates")

            except Exception as export_e:
                print(f"Error exporting data: {str(export_e)}")
                # Continue with repair even if export fails

            # 3. Close connections and recreate the database file
            self.close()

            # Remove the corrupted database
            if os.path.exists(self.db_path):
                os.remove(self.db_path)
                print(f"Removed corrupted database: {self.db_path}")

            # Create a fresh database
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            self.create_tables()
            print("Created fresh database")

            # 4. Import the templates back
            import_count = 0

            # Create a temporary connection to avoid conflicts
            temp_conn = sqlite3.connect(self.db_path)
            temp_cursor = temp_conn.cursor()

            for template in templates:
                try:
                    # Convert Python objects to JSON strings
                    regions_json = json.dumps(template['regions'])
                    column_lines_json = json.dumps(template['column_lines'])
                    config_json = json.dumps(template['config'])

                    # Handle page-specific data for multi-page templates
                    page_regions_json = json.dumps(template['page_regions']) if template.get('page_regions') else None
                    page_column_lines_json = json.dumps(template['page_column_lines']) if template.get('page_column_lines') else None
                    validation_rules_json = json.dumps(template['validation_rules']) if template.get('validation_rules') else None
                    json_template_json = json.dumps(template['json_template']) if template.get('json_template') else None

                    # Insert the template
                    temp_cursor.execute("""
                        INSERT INTO templates
                        (id, name, description, regions, column_lines, config, template_type,
                         creation_date, last_modified, page_count, page_regions,
                         page_column_lines, validation_rules, json_template)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        template['id'], template['name'], template['description'],
                        regions_json, column_lines_json, config_json,
                        template['template_type'], template['creation_date'],
                        template['last_modified'], template['page_count'],
                        page_regions_json, page_column_lines_json, validation_rules_json, json_template_json
                    ))

                    import_count += 1
                except Exception as import_e:
                    print(f"Error importing template {template.get('name', 'unknown')}: {str(import_e)}")

                temp_conn.commit()
                print(f"Successfully imported {import_count} of {exported_count} templates")

            # Close the temporary connection
            temp_conn.close()

            # Check disk space before running optimization
            has_space, space_message = check_disk_space(self.db_path)
            print(space_message)

            if has_space:
                # Run optimization after repair
                print("Running optimization after repair...")
                self.optimize_database()
            else:
                print("Skipping optimization after repair due to insufficient disk space.")

            if exported_count > 0 and import_count == exported_count:
                return True, f"Database successfully repaired: recovered {import_count} templates"
            elif import_count > 0:
                return True, f"Database partially repaired: recovered {import_count} of {exported_count} templates"
            else:
                return False, "Database structure repaired but no data could be recovered"

        except Exception as e:
            error_msg = f"Error during database repair: {str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            return False, error_msg




