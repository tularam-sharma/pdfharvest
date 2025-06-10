#!/usr/bin/env python3
"""
Database migration script to add extraction_method column to templates table

This script adds the extraction_method column to the existing templates table
to support multiple PDF extraction methods (pypdf_table_extraction, pdftotext, tesseract_ocr).
"""

import os
import sqlite3
import sys
from datetime import datetime

def check_database_exists():
    """Check if the database file exists"""
    if not os.path.exists("invoice_templates.db"):
        print("❌ Database file 'invoice_templates.db' not found.")
        print("   Please run the main application first to create the database.")
        return False
    return True

def check_column_exists(cursor):
    """Check if extraction_method column already exists"""
    cursor.execute("PRAGMA table_info(templates)")
    columns = [column[1] for column in cursor.fetchall()]
    return 'extraction_method' in columns

def backup_database():
    """Create a backup of the database before migration"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"invoice_templates_backup_{timestamp}.db"
    
    try:
        import shutil
        shutil.copy2("invoice_templates.db", backup_name)
        print(f"✅ Database backed up to: {backup_name}")
        return backup_name
    except Exception as e:
        print(f"⚠️  Warning: Could not create backup: {e}")
        return None

def add_extraction_method_column(cursor):
    """Add the extraction_method column to the templates table"""
    try:
        # Add the column with default value
        cursor.execute("""
            ALTER TABLE templates 
            ADD COLUMN extraction_method TEXT DEFAULT 'pypdf_table_extraction'
        """)
        print("✅ Added extraction_method column to templates table")
        return True
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("ℹ️  extraction_method column already exists")
            return True
        else:
            print(f"❌ Error adding column: {e}")
            return False

def update_existing_templates(cursor):
    """Update existing templates to have the default extraction method"""
    try:
        # Count existing templates
        cursor.execute("SELECT COUNT(*) FROM templates")
        count = cursor.fetchone()[0]
        
        if count > 0:
            # Update templates that have NULL extraction_method
            cursor.execute("""
                UPDATE templates 
                SET extraction_method = 'pypdf_table_extraction' 
                WHERE extraction_method IS NULL
            """)
            
            updated = cursor.rowcount
            print(f"✅ Updated {updated} existing templates with default extraction method")
        else:
            print("ℹ️  No existing templates to update")
        
        return True
    except Exception as e:
        print(f"❌ Error updating existing templates: {e}")
        return False

def verify_migration(cursor):
    """Verify that the migration was successful"""
    try:
        # Check that column exists
        cursor.execute("PRAGMA table_info(templates)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'extraction_method' not in columns:
            print("❌ Migration verification failed: column not found")
            return False
        
        # Check that we can insert and retrieve data
        test_data = {
            'name': 'migration_test_template',
            'description': 'Test template for migration verification',
            'template_type': 'single',
            'extraction_method': 'pdftotext',
            'regions': '{}',
            'column_lines': '{}',
            'config': '{}'
        }
        
        # Insert test template
        cursor.execute("""
            INSERT INTO templates (name, description, template_type, extraction_method, regions, column_lines, config)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            test_data['name'],
            test_data['description'],
            test_data['template_type'],
            test_data['extraction_method'],
            test_data['regions'],
            test_data['column_lines'],
            test_data['config']
        ))
        
        # Retrieve test template
        cursor.execute("SELECT extraction_method FROM templates WHERE name = ?", (test_data['name'],))
        result = cursor.fetchone()
        
        if result and result[0] == 'pdftotext':
            print("✅ Migration verification successful")
            
            # Clean up test template
            cursor.execute("DELETE FROM templates WHERE name = ?", (test_data['name'],))
            return True
        else:
            print("❌ Migration verification failed: could not retrieve test data")
            return False
            
    except Exception as e:
        print(f"❌ Migration verification failed: {e}")
        return False

def show_extraction_methods():
    """Show available extraction methods"""
    print("\n📋 Available extraction methods:")
    print("   • pypdf_table_extraction (default) - Standard table extraction using pypdf")
    print("   • pdftotext - Text extraction using invoice2data's pdftotext parser")
    print("   • tesseract_ocr - OCR extraction using invoice2data's tesseract parser")
    print()

def main():
    """Run the migration"""
    print("=" * 70)
    print("Database Migration: Adding extraction_method Column")
    print("=" * 70)
    print()
    
    # Check if database exists
    if not check_database_exists():
        return False
    
    # Create backup
    backup_name = backup_database()
    
    try:
        # Connect to database
        conn = sqlite3.connect("invoice_templates.db")
        cursor = conn.cursor()
        
        # Check if column already exists
        if check_column_exists(cursor):
            print("ℹ️  extraction_method column already exists in templates table")
            print("✅ Migration not needed")
            conn.close()
            show_extraction_methods()
            return True
        
        print("🔄 Starting migration...")
        
        # Add the column
        if not add_extraction_method_column(cursor):
            conn.close()
            return False
        
        # Update existing templates
        if not update_existing_templates(cursor):
            conn.close()
            return False
        
        # Commit changes
        conn.commit()
        print("✅ Migration changes committed")
        
        # Verify migration
        if not verify_migration(cursor):
            conn.close()
            return False
        
        # Final commit
        conn.commit()
        conn.close()
        
        print("\n🎉 Migration completed successfully!")
        show_extraction_methods()
        
        print("💡 Next steps:")
        print("   1. Templates can now specify their preferred extraction method")
        print("   2. Use the template manager to set extraction methods for templates")
        print("   3. The system will automatically use the specified method during extraction")
        print()
        
        return True
        
    except Exception as e:
        print(f"❌ Migration failed: {e}")
        if 'conn' in locals():
            conn.close()
        
        if backup_name:
            print(f"💾 Database backup available at: {backup_name}")
        
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
