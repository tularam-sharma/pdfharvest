"""
Invoice Processing Utilities (Simplified)

This module provides common utilities for invoice processing, including:
- Middle page logic for multi-page invoices
- Fixed page logic for multi-page invoices
- Invoice2data processing functions
- PDF extraction functions

It centralizes the processing logic to reduce code duplication between
split_screen_invoice_processor.py and bulk_processor.py.
"""

import os
import json
import pandas as pd
import fitz  # PyMuPDF
import re
from typing import Dict, List, Tuple, Union, Optional, Any

# Import utility modules
from pdf_extraction_utils import (
    extract_table, extract_tables, clean_dataframe, DEFAULT_EXTRACTION_PARAMS,
    clear_extraction_cache, clear_extraction_cache_for_pdf, clear_extraction_cache_for_section,
    get_scale_factors, convert_display_to_pdf_coords, convert_pdf_to_display_coords, get_extraction_cache_stats
)
from multi_method_extraction import extract_with_method, cleanup_extraction

# Check if invoice2data is available
try:
    import invoice2data
    from invoice2data import extract_data
    from invoice2data.extract.loader import read_templates
    INVOICE2DATA_AVAILABLE = True
except ImportError:
    INVOICE2DATA_AVAILABLE = False
    print("Warning: invoice2data module not available. Some functionality will be limited.")

def get_template_page_for_pdf_page(pdf_page_index: int, pdf_total_pages: int, template_data: Dict) -> int:
    """Determine which template page to use for a given PDF page

    Args:
        pdf_page_index (int): The 0-based index of the PDF page
        pdf_total_pages (int): The total number of pages in the PDF
        template_data (dict): The template data dictionary

    Returns:
        int: The 0-based index of the template page to use
    """
    # Get template type and page count
    template_type = template_data.get('template_type', 'single')
    template_page_count = template_data.get('page_count', 1)

    # For single-page templates, always use the first (and only) page
    if template_type == 'single':
        return 0

    # Handle edge case: no template pages available
    if template_page_count <= 0:
        return 0

    # Complex page mapping logic removed - using simplified page-wise approach
    # Page mapping features will be implemented later as a separate enhancement

    # For now, use simple page mapping: use the exact page index, but don't exceed template page count
    return min(pdf_page_index, template_page_count - 1)


# Complex template application function removed - using simplified page-wise approach
# Page mapping features will be implemented later as a separate enhancement

def apply_template_with_middle_page_logic(template_data: Dict, pdf_path: str, pdf_total_pages: int) -> Dict:
    """Apply a template to a PDF with simplified page-wise logic

    Simplified version that applies template pages directly to PDF pages.
    Complex page mapping logic has been removed.

    Args:
        template_data (dict): The template data dictionary
        pdf_path (str): Path to the PDF file
        pdf_total_pages (int): Total number of pages in the PDF

    Returns:
        dict: A dictionary mapping PDF page indices to template page indices and regions/columns
    """
    result = {}

    # For each page in the PDF, determine which template page to use
    for pdf_page_idx in range(pdf_total_pages):
        template_page_idx = get_template_page_for_pdf_page(pdf_page_idx, pdf_total_pages, template_data)

        # Get regions and column lines for this template page
        regions = {}
        column_lines = {}

        if template_data.get('template_type') == 'single':
            # For single-page templates, use the regions and column_lines directly
            regions = template_data.get('regions', {})
            column_lines = template_data.get('column_lines', {})
        else:
            # For multi-page templates, get the regions and column_lines for the specific page
            page_regions = template_data.get('page_regions', [])
            page_column_lines = template_data.get('page_column_lines', [])

            if template_page_idx < len(page_regions):
                regions = page_regions[template_page_idx]

            if template_page_idx < len(page_column_lines):
                column_lines = page_column_lines[template_page_idx]

        # Store the mapping
        result[pdf_page_idx] = {
            'template_page_idx': template_page_idx,
            'regions': regions,
            'column_lines': column_lines
        }

    return result


def extract_multi_page_invoice(pdf_path: str, template_data: Dict, extraction_params: Dict = None) -> Dict:
    """Extract data from a multi-page invoice using cached results when available
    
    Args:
        pdf_path (str): Path to the PDF file
        template_data (dict): Template data from the database
        extraction_params (dict, optional): Custom extraction parameters
        
    Returns:
        dict: Dictionary containing extracted data for all pages
    """
    # Initialize data structure
    all_data = {'header': None, 'items': [], 'summary': None}
    
    try:
        # Get PDF page count
        with fitz.open(pdf_path) as pdf:
            pdf_total_pages = len(pdf)
            
        # Apply template with middle page logic
        page_mapping = apply_template_with_middle_page_logic(
            template_data=template_data,
            pdf_path=pdf_path,
            pdf_total_pages=pdf_total_pages
        )
        
        # Process each page
        for pdf_page_idx in range(pdf_total_pages):
            # Get mapping for this page
            mapping = page_mapping.get(pdf_page_idx, {})
            regions = mapping.get('regions', {})
            column_lines = mapping.get('column_lines', {})
            
            # Skip if no regions defined for this page
            if not regions:
                continue
                
            # Extract tables for this page
            page_data = extract_invoice_tables(
                pdf_path=pdf_path,
                regions=regions,
                column_lines=column_lines,
                extraction_params=extraction_params,
                page_number=pdf_page_idx + 1,  # 1-based page number
                use_cache=True
            )
            
            # Process each section
            for section in ['header', 'items', 'summary']:
                if section in page_data and page_data[section]:
                    for i, df in enumerate(page_data[section]):
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            # Ensure region_label includes page number
                            if 'region_label' not in df.columns:
                                prefix = {'header': 'H', 'items': 'I', 'summary': 'S'}[section]
                                df['region_label'] = f"{prefix}{i+1}_R1_P{pdf_page_idx + 1}"
                            
                            # Add page_number column if not present
                            if 'page_number' not in df.columns:
                                df['page_number'] = pdf_page_idx + 1
                            
                            # Process based on section type
                            if section == 'header':
                                if all_data['header'] is None:
                                    all_data['header'] = df
                                    print(f"[DEBUG] Initialized header with data from page {pdf_page_idx + 1}")
                                else:
                                    # Use merge instead of concat to preserve all data
                                    # Outer join ensures all rows from both DataFrames are kept
                                    all_data['header'] = pd.merge(
                                        all_data['header'], 
                                        df,
                                        on='region_label',  # Merge on region_label
                                        how='outer',        # Keep all rows from both DataFrames
                                        indicator=True      # Add _merge column to track source
                                    )
                                    print(f"[DEBUG] Merged header data from page {pdf_page_idx + 1} using region_label")
                            elif section == 'items':
                                all_data['items'].append(df)
                                print(f"[DEBUG] Added items data from page {pdf_page_idx + 1}")
                            elif section == 'summary':
                                if all_data['summary'] is None:
                                    all_data['summary'] = df
                                    print(f"[DEBUG] Initialized summary with data from page {pdf_page_idx + 1}")
                                else:
                                    # Use merge instead of concat to preserve all data
                                    all_data['summary'] = pd.merge(
                                        all_data['summary'], 
                                        df,
                                        on='region_label',  # Merge on region_label
                                        how='outer',        # Keep all rows from both DataFrames
                                        indicator=True      # Add _merge column to track source
                                    )
                                    print(f"[DEBUG] Merged summary data from page {pdf_page_idx + 1} using region_label")
        
        # Clean up indicator columns if present
        for section in ['header', 'summary']:
            if isinstance(all_data[section], pd.DataFrame) and '_merge' in all_data[section].columns:
                all_data[section] = all_data[section].drop(columns=['_merge'])
                
        return all_data
        
    except Exception as e:
        print(f"Error in extract_multi_page_invoice: {str(e)}")
        import traceback
        traceback.print_exc()
        return {'header': None, 'items': pd.DataFrame(), 'summary': None}


def extract_invoice_tables(
    pdf_path: str,
    regions: Dict[str, List[List[float]]],
    column_lines: Dict[str, List[List[float]]] = None,
    extraction_params: Dict = None,
    page_number: int = 1,
    use_cache: bool = True
) -> Dict[str, List[pd.DataFrame]]:
    """Extract tables from a PDF invoice

    Args:
        pdf_path (str): Path to the PDF file
        regions (dict): Dictionary of regions by section type
        column_lines (dict, optional): Dictionary of column lines by section type
        extraction_params (dict, optional): Custom extraction parameters
        page_number (int, optional): Page number to extract from (1-based)
        use_cache (bool, optional): Whether to use the extraction cache

    Returns:
        dict: Dictionary of extracted tables by section type
    """
    # Initialize result
    result = {
        'header': [],
        'items': [],
        'summary': [],
        'extraction_status': {
            'header': 'not_attempted',
            'items': 'not_attempted',
            'summary': 'not_attempted'
        }
    }

    # Use default extraction parameters if none provided
    if extraction_params is None:
        extraction_params = DEFAULT_EXTRACTION_PARAMS

    # Initialize column_lines if not provided
    if column_lines is None:
        column_lines = {}

    # Process each section
    for section in ['header', 'items', 'summary']:
        if section in regions and regions[section]:
            section_regions = regions[section]
            section_columns = column_lines.get(section, [])

            # Use multi-method extraction with default method
            extraction_method = "pypdf_table_extraction"  # Default method
            tables = extract_with_method(
                pdf_path=pdf_path,
                extraction_method=extraction_method,
                page_number=page_number,
                table_areas=section_regions,
                columns_list=section_columns,
                section_type=section,
                extraction_params=extraction_params,
                use_cache=use_cache
            )

            if tables:
                result[section] = tables
                result['extraction_status'][section] = 'success'
            else:
                result['extraction_status'][section] = 'failed'

    return result


def convert_extraction_to_text(extraction_data: Dict, pdf_path: str = None) -> str:
    """Convert the extraction data to a text format that can be used by invoice2data

    In multi-page PDF extraction mode, we preserve region data from all pages without duplicate checking,
    as the same region can appear on different pages.

    Args:
        extraction_data (dict): The extracted data from pypdf_table_extraction
        pdf_path (str, optional): Path to the PDF file, used for adding metadata if not present

    Returns:
        str: Text representation of the extracted data with pipe-separated values
    """
    if not extraction_data:
        return ""

    # Make a copy of the extraction data to avoid modifying the original
    extraction_data = extraction_data.copy()

    # Add metadata if not already present
    if 'metadata' not in extraction_data:
        extraction_data['metadata'] = {}

    # Add PDF filename to metadata if not already present
    if pdf_path and 'filename' not in extraction_data['metadata']:
        extraction_data['metadata']['filename'] = os.path.basename(pdf_path)

    # Check if we have template_type in metadata
    if 'template_type' in extraction_data['metadata']:
        template_type = extraction_data['metadata']['template_type']
        print(f"[DEBUG] Template type: {template_type}")

        # If this is a multi-page template, make sure we're showing data from all pages
        if template_type == 'multi':
            print(f"[DEBUG] Multi-page template detected, ensuring data from all pages is shown")

    # Initialize text lines
    text_lines = []

    # Add metadata section
    text_lines.append("METADATA")
    for key, value in extraction_data['metadata'].items():
        text_lines.append(f"{key}|{value}")
    text_lines.append("")  # Empty line to separate sections

    # Process header section
    text_lines.append("HEADER")
    if 'header' in extraction_data and extraction_data['header'] is not None:
        header_data = extraction_data['header']

        # Handle list of DataFrames (multiple regions or pages)
        if isinstance(header_data, list):
            print(f"[DEBUG] Processing header data as list of {len(header_data)} DataFrames")
            for i, df in enumerate(header_data):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Get page number for this DataFrame if available
                    page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else i+1
                    print(f"[DEBUG] Processing header DataFrame {i+1} from page {page_num} with {len(df)} rows")

                    # Add region labels to the text
                    for _, row in df.iterrows():
                        if 'region_label' in row.index and pd.notna(row['region_label']):
                            # Use the region_label which should already include page number
                            label = row['region_label']
                            values = []
                            # Add all other columns except region_label, page_number, etc.
                            for col, val in row.items():
                                if col not in ['region_label', 'page_number', '_page_number', '_row_number', 'pdf_page']:
                                    if pd.notna(val):
                                        values.append(str(val))
                            if values:
                                text_lines.append(f"{label}|{'|'.join(values)}")
                                print(f"[DEBUG] Added header line: {label}|{'|'.join(values)}")
        # Handle single DataFrame
        elif isinstance(header_data, pd.DataFrame) and not header_data.empty:
            print(f"[DEBUG] Processing header data as single DataFrame with {len(header_data)} rows")

            # Add region labels to the text
            for _, row in header_data.iterrows():
                if 'region_label' in row.index and pd.notna(row['region_label']):
                    # Use the region_label which should already include page number
                    label = row['region_label']
                    values = []
                    # Add all other columns except region_label, page_number, etc.
                    for col, val in row.items():
                        if col not in ['region_label', 'page_number', '_page_number', '_row_number', 'pdf_page']:
                            if pd.notna(val):
                                values.append(str(val))
                    if values:
                        text_lines.append(f"{label}|{'|'.join(values)}")
                        print(f"[DEBUG] Added header line: {label}|{'|'.join(values)}")
    text_lines.append("")  # Empty line to separate sections

    # Process items section
    text_lines.append("ITEMS")
    if 'items' in extraction_data and extraction_data['items'] is not None:
        items_data = extraction_data['items']

        # Handle list of DataFrames (multiple regions or pages)
        if isinstance(items_data, list):
            print(f"[DEBUG] Processing items data as list of {len(items_data)} DataFrames")
            for i, df in enumerate(items_data):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Get page number for this DataFrame if available
                    page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else i+1
                    print(f"[DEBUG] Processing items DataFrame {i+1} from page {page_num} with {len(df)} rows")

                    # Add region labels to the text
                    for _, row in df.iterrows():
                        if 'region_label' in row.index and pd.notna(row['region_label']):
                            # Use the region_label which should already include page number
                            label = row['region_label']
                            values = []
                            # Add all other columns except region_label, page_number, etc.
                            for col, val in row.items():
                                if col not in ['region_label', 'page_number', '_page_number', '_row_number', 'pdf_page']:
                                    if pd.notna(val):
                                        values.append(str(val))
                            if values:
                                text_lines.append(f"{label}|{'|'.join(values)}")
                                print(f"[DEBUG] Added items line: {label}|{'|'.join(values[:1])}...")
        # Handle single DataFrame
        elif isinstance(items_data, pd.DataFrame) and not items_data.empty:
            print(f"[DEBUG] Processing items data as single DataFrame with {len(items_data)} rows")

            # Add region labels to the text
            for _, row in items_data.iterrows():
                if 'region_label' in row.index and pd.notna(row['region_label']):
                    # Use the region_label which should already include page number
                    label = row['region_label']
                    values = []
                    # Add all other columns except region_label, page_number, etc.
                    for col, val in row.items():
                        if col not in ['region_label', 'page_number', '_page_number', '_row_number', 'pdf_page']:
                            if pd.notna(val):
                                values.append(str(val))
                    if values:
                        text_lines.append(f"{label}|{'|'.join(values)}")
                        print(f"[DEBUG] Added items line: {label}|{'|'.join(values[:1])}...")
    text_lines.append("")  # Empty line to separate sections

    # Process summary section
    text_lines.append("SUMMARY")
    if 'summary' in extraction_data and extraction_data['summary'] is not None:
        summary_data = extraction_data['summary']

        # Handle list of DataFrames (multiple regions or pages)
        if isinstance(summary_data, list):
            print(f"[DEBUG] Processing summary data as list of {len(summary_data)} DataFrames")
            for i, df in enumerate(summary_data):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Get page number for this DataFrame if available
                    page_num = df['page_number'].iloc[0] if 'page_number' in df.columns else i+1
                    print(f"[DEBUG] Processing summary DataFrame {i+1} from page {page_num} with {len(df)} rows")

                    # Add region labels to the text
                    for _, row in df.iterrows():
                        if 'region_label' in row.index and pd.notna(row['region_label']):
                            # Use the region_label which should already include page number
                            label = row['region_label']
                            values = []
                            # Add all other columns except region_label, page_number, etc.
                            for col, val in row.items():
                                if col not in ['region_label', 'page_number', '_page_number', '_row_number', 'pdf_page']:
                                    if pd.notna(val):
                                        values.append(str(val))
                            if values:
                                text_lines.append(f"{label}|{'|'.join(values)}")
                                print(f"[DEBUG] Added summary line: {label}|{'|'.join(values[:1])}...")
        # Handle single DataFrame
        elif isinstance(summary_data, pd.DataFrame) and not summary_data.empty:
            print(f"[DEBUG] Processing summary data as single DataFrame with {len(summary_data)} rows")

            # Add region labels to the text
            for _, row in summary_data.iterrows():
                if 'region_label' in row.index and pd.notna(row['region_label']):
                    # Use the region_label which should already include page number
                    label = row['region_label']
                    values = []
                    # Add all other columns except region_label, page_number, etc.
                    for col, val in row.items():
                        if col not in ['region_label', 'page_number', '_page_number', '_row_number', 'pdf_page']:
                            if pd.notna(val):
                                values.append(str(val))
                    if values:
                        text_lines.append(f"{label}|{'|'.join(values)}")
                        print(f"[DEBUG] Added summary line: {label}|{'|'.join(values[:1])}...")

    # Print the generated text for debugging
    print(f"[DEBUG] Generated extraction text ({len(text_lines)} lines):")
    for i, line in enumerate(text_lines[:20]):  # Print first 20 lines
        print(f"[DEBUG] Line {i+1}: {line}")
    if len(text_lines) > 20:
        print(f"[DEBUG] ... and {len(text_lines) - 20} more lines")

    # Join all lines with newlines
    return "\n".join(text_lines)


def build_invoice2data_template(issuer, fields_data, options_data=None, keywords=None, exclude_keywords=None, lines_data=None):
    """Build the invoice2data template dictionary

    Args:
        issuer (str): The issuer name
        fields_data (dict): Dictionary of field names and their regex patterns/types
        options_data (dict, optional): Dictionary of options like currency, languages, etc.
        keywords (list, optional): List of keywords for template matching
        exclude_keywords (list, optional): List of keywords to exclude
        lines_data (dict, optional): Dictionary of line-related settings

    Returns:
        dict: The invoice2data template dictionary
    """
    # Create the template dictionary
    template = {
        "issuer": issuer,
        "fields": fields_data,
        "keywords": keywords or [],
        "exclude_keywords": exclude_keywords or [],
        "options": options_data or {}
    }

    # Add lines data if provided
    if lines_data:
        template["lines"] = lines_data

    return template


def process_with_invoice2data(pdf_path: str, template_data: Dict, extracted_data: Dict = None, temp_dir: str = None) -> Dict:
    """Process a PDF file with invoice2data using the template and extracted data

    Simplified version that focuses only on the core functionality.

    Args:
        pdf_path (str): Path to the PDF file
        template_data (dict): Template data from the database
        extracted_data (dict, optional): Extracted data from the PDF
        temp_dir (str, optional): Path to temporary directory

    Returns:
        dict: The extraction result from invoice2data, or None if extraction failed
    """
    if not INVOICE2DATA_AVAILABLE:
        print("invoice2data module is not available. Skipping invoice2data processing.")
        return None

    # Initialize variables
    result = None
    cleanup_temp_dir = False
    import tempfile

    # Create a temporary directory if not provided
    if not temp_dir:
        temp_dir = tempfile.mkdtemp(prefix="invoice2data_")
        print(f"Created temporary directory: {temp_dir}")
        cleanup_temp_dir = True
    else:
        cleanup_temp_dir = False

    try:
        # Get the template from the template_data
        template = template_data.get('json_template')
        if not template:
            print("No template found in template_data")
            return None

        # Get a safe template name for filenames
        safe_template_name = "template"
        if 'issuer' in template and template['issuer']:
            safe_template_name = template['issuer'].replace(' ', '_').replace('/', '_').replace('\\', '_')

        # Save the template to a temporary file
        temp_template_path = os.path.join(temp_dir, f"{safe_template_name}.yml")
        with open(temp_template_path, "w", encoding="utf-8") as f:
            import yaml
            yaml.safe_dump(
                template,
                f,
                default_flow_style=False,
                allow_unicode=True,
                sort_keys=False,
                indent=2,
                width=80
            )

        # Convert the extracted data to text format for invoice2data
        if extracted_data:
            text_content = convert_extraction_to_text(extracted_data, pdf_path=pdf_path)

            # Save the text to a temporary file
            temp_text_path = os.path.join(temp_dir, f"{safe_template_name}.txt")
            with open(temp_text_path, "w", encoding="utf-8") as f:
                f.write(text_content)
        else:
            # If no extracted data provided, use the PDF directly
            temp_text_path = pdf_path

        # Try to use the invoice2data API directly
        try:
            # Load the template
            templates = read_templates(temp_dir)

            # Extract data
            result = extract_data(temp_text_path, templates=templates)
            return result

        except Exception as api_error:
            print(f"Error using invoice2data API: {str(api_error)}")
            return None

    except Exception as e:
        print(f"Error in process_with_invoice2data: {str(e)}")
        result = None
    finally:
        # Clean up temporary directory if we created it
        if cleanup_temp_dir and os.path.exists(temp_dir):
            import shutil
            shutil.rmtree(temp_dir)

    return result


def analyze_invoice2data_warnings(warnings_text, template):
    """Analyze invoice2data warnings and suggest fixes

    Args:
        warnings_text (str): The warnings captured from invoice2data
        template (dict): The template that was used

    Returns:
        str: Suggestions for fixing the warnings
    """
    suggestions = []

    # Check for common warning patterns
    if "No keywords found" in warnings_text or "keywords" in warnings_text.lower():
        suggestions.append("Keywords issue detected:")
        suggestions.append("- Check that your keywords match text in the invoice")
        suggestions.append("- Add more specific keywords that appear in the invoice")
        suggestions.append("- Make sure keywords are not too generic")

    if "No matching template" in warnings_text:
        suggestions.append("Template matching issue detected:")
        suggestions.append("- Verify that the template keywords match text in the invoice")
        suggestions.append("- Check if exclude_keywords might be preventing a match")

    if "regex" in warnings_text.lower() or "pattern" in warnings_text.lower():
        suggestions.append("Regex pattern issue detected:")
        suggestions.append("- Check that your regex patterns have capturing groups (parentheses)")
        suggestions.append("- Verify that the patterns match the format in the invoice")
        suggestions.append("- Test your regex patterns with a tool like regex101.com")

    # If no specific warnings were found, provide general suggestions
    if not suggestions:
        # Analyze the template for potential issues
        if template:
            if not template.get("keywords") or len(template.get("keywords", [])) == 0:
                suggestions.append("Template has no keywords defined")
                suggestions.append("- Add keywords that appear in the invoice")

            if not template.get("fields") or len(template.get("fields", {})) == 0:
                suggestions.append("Template has no fields defined")
                suggestions.append("- Add fields with regex patterns to extract data")

            # Check for required fields
            required_fields = ["invoice_number", "date", "amount"]
            missing_fields = [field for field in required_fields if field not in template.get("fields", {})]
            if missing_fields:
                suggestions.append(f"Template is missing required fields: {', '.join(missing_fields)}")
                suggestions.append("- Add these fields with appropriate regex patterns")

        # Add general suggestions if no specific issues were found
        if not suggestions:
            suggestions.append("No specific issues detected. General suggestions:")
            suggestions.append("- Verify that the template keywords match text in the invoice")
            suggestions.append("- Check that regex patterns have capturing groups (parentheses)")
            suggestions.append("- Ensure date formats match those in the invoice")
            suggestions.append("- Make sure the field names match those in the invoice")

    return "\n".join(suggestions)


