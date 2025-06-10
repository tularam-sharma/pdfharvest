#!/usr/bin/env python3
"""
Create the invoice templates database with extraction_method support

This script creates a new database with the extraction_method column included
from the beginning, so no migration is needed.
"""

import sqlite3
import os
import sys

def create_database():
    """Create the database with all required tables and columns"""
    try:
        # Remove existing database if it exists
        if os.path.exists("invoice_templates.db"):
            print("Removing existing database...")
            os.remove("invoice_templates.db")
        
        # Create new database
        conn = sqlite3.connect("invoice_templates.db")
        cursor = conn.cursor()
        
        # Create templates table with extraction_method column
        cursor.execute("""
            CREATE TABLE templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT,
                template_type TEXT NOT NULL DEFAULT 'single',
                regions TEXT,
                column_lines TEXT,
                config TEXT,
                creation_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                page_count INTEGER DEFAULT 1,
                page_regions TEXT,
                page_column_lines TEXT,
                page_configs TEXT,
                json_template TEXT,
                drawing_regions TEXT,
                drawing_column_lines TEXT,
                extraction_regions TEXT,
                extraction_column_lines TEXT,
                drawing_page_regions TEXT,
                drawing_page_column_lines TEXT,
                extraction_page_regions TEXT,
                extraction_page_column_lines TEXT,
                extraction_method TEXT DEFAULT 'pypdf_table_extraction'
            )
        """)
        
        print("✅ Created templates table with extraction_method column")
        
        # Create any other tables that might be needed
        # (Add more table creation statements here if needed)
        
        # Commit and close
        conn.commit()
        conn.close()
        
        print("✅ Database created successfully with extraction method support")
        return True
        
    except Exception as e:
        print(f"❌ Error creating database: {e}")
        return False

def test_database():
    """Test that the database was created correctly"""
    try:
        conn = sqlite3.connect("invoice_templates.db")
        cursor = conn.cursor()
        
        # Check table structure
        cursor.execute("PRAGMA table_info(templates)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'extraction_method' in columns:
            print("✅ extraction_method column found in templates table")
        else:
            print("❌ extraction_method column not found")
            return False
        
        # Test inserting a template with extraction method
        test_template = {
            'name': 'test_template',
            'description': 'Test template with extraction method',
            'template_type': 'single',
            'extraction_method': 'pdftotext',
            'regions': '{"header": [], "items": [], "summary": []}',
            'column_lines': '{"header": [], "items": [], "summary": []}',
            'config': '{"header": {"row_tol": 5}, "items": {"row_tol": 15}, "summary": {"row_tol": 10}}'
        }
        
        cursor.execute("""
            INSERT INTO templates (name, description, template_type, extraction_method, regions, column_lines, config)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            test_template['name'],
            test_template['description'],
            test_template['template_type'],
            test_template['extraction_method'],
            test_template['regions'],
            test_template['column_lines'],
            test_template['config']
        ))
        
        # Retrieve the template
        cursor.execute("SELECT extraction_method FROM templates WHERE name = ?", (test_template['name'],))
        result = cursor.fetchone()
        
        if result and result[0] == 'pdftotext':
            print("✅ Successfully inserted and retrieved template with extraction method")
        else:
            print("❌ Failed to retrieve correct extraction method")
            return False
        
        # Clean up test template
        cursor.execute("DELETE FROM templates WHERE name = ?", (test_template['name'],))
        conn.commit()
        conn.close()
        
        print("✅ Database test completed successfully")
        return True
        
    except Exception as e:
        print(f"❌ Database test failed: {e}")
        return False

def main():
    """Create and test the database"""
    print("=" * 60)
    print("Creating Invoice Templates Database with Extraction Method Support")
    print("=" * 60)
    print()
    
    # Create database
    if not create_database():
        return False
    
    # Test database
    if not test_database():
        return False
    
    print("\n🎉 Database creation completed successfully!")
    print("\n📋 Available extraction methods:")
    print("   • pypdf_table_extraction (default) - Standard table extraction using pypdf")
    print("   • pdftotext - Text extraction using invoice2data's pdftotext parser")
    print("   • tesseract_ocr - OCR extraction using invoice2data's tesseract parser")
    print()
    print("💡 The database is now ready for use with multi-method extraction support.")
    print()
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
