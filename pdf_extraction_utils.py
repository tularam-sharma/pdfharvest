"""
PDF Extraction Utilities

This module provides common utilities for PDF table extraction using pypdf_table_extraction.
It centralizes the extraction logic to reduce code duplication and complexity.
"""

import fitz  # PyMuPDF
import pypdf_table_extraction
import pandas as pd
import numpy as np
import os
import re
import hashlib
from typing import Dict, List, Tuple, Union, Optional, Any

# Import standardized error handling
from error_handler import log_error, log_info, log_warning, handle_exception, ErrorContext

# Global extraction cache to avoid redundant extractions
_EXTRACTION_CACHE = {}

# Global cache for multi-page extraction results
_MULTIPAGE_CACHE = {}

def _get_cache_key(pdf_path: str, page_number: int, table_area: Union[List[float], str], columns: Union[List[float], str], section_type: str, region_label: str = None) -> str:
    """Generate a unique cache key for extraction parameters

    Args:
        pdf_path: Path to the PDF file
        page_number: Page number (1-based)
        table_area: Table area coordinates or string
        columns: Column coordinates or string
        section_type: Section type ('header', 'items', or 'summary')
        region_label: Optional region label for more specific caching

    Returns:
        str: MD5 hash of the parameters to use as a cache key
    """
    # Convert table_area to string if it's a list
    if isinstance(table_area, list):
        table_area_str = ','.join(map(str, table_area))
    else:
        table_area_str = str(table_area) if table_area else 'None'

    # Convert columns to string if it's a list
    if isinstance(columns, list):
        columns_str = ','.join(map(str, columns))
    else:
        columns_str = str(columns) if columns else 'None'

    # Include region_label in the key if provided
    region_part = f"|{region_label}" if region_label else ""

    # Create a string with all parameters
    key_str = f"{pdf_path}|{page_number}|{table_area_str}|{columns_str}|{section_type}{region_part}"

    # Hash the string to create a shorter key
    return hashlib.md5(key_str.encode()).hexdigest()

def get_multipage_cache_key(pdf_path: str) -> str:
    """Generate a unique cache key for multi-page extraction results

    Args:
        pdf_path: Path to the PDF file

    Returns:
        str: MD5 hash of the PDF path to use as a cache key
    """
    # Use just the PDF path for multi-page cache
    return hashlib.md5(pdf_path.encode()).hexdigest()

def clear_extraction_cache():
    """Clear the global extraction cache"""
    global _EXTRACTION_CACHE, _MULTIPAGE_CACHE
    _EXTRACTION_CACHE = {}
    _MULTIPAGE_CACHE = {}
    print(f"[DEBUG] Extraction cache and multi-page cache cleared")

def get_extraction_cache_stats():
    """Get statistics about the extraction cache"""
    return {
        'extraction_cache_size': len(_EXTRACTION_CACHE),
        'extraction_cache_keys': list(_EXTRACTION_CACHE.keys())[:10],  # Return first 10 keys for debugging
        'multipage_cache_size': len(_MULTIPAGE_CACHE),
        'multipage_cache_keys': list(_MULTIPAGE_CACHE.keys())[:10]  # Return first 10 keys for debugging
    }

def clear_extraction_cache_for_pdf(pdf_path: str, preserve_multipage: bool = False) -> None:
    """Clear the extraction cache for a specific PDF file

    Args:
        pdf_path: Path to the PDF file
        preserve_multipage: Whether to preserve the multi-page cache (default: False)
    """
    global _EXTRACTION_CACHE, _MULTIPAGE_CACHE

    # Clear individual extraction cache entries
    pdf_hash_prefix = hashlib.md5(pdf_path.encode()).hexdigest()[:8]
    keys_to_remove = [k for k in _EXTRACTION_CACHE.keys() if k.startswith(pdf_hash_prefix)]
    for key in keys_to_remove:
        del _EXTRACTION_CACHE[key]

    # Clear multi-page cache entry only if not preserving it
    if not preserve_multipage:
        multipage_key = get_multipage_cache_key(pdf_path)
        if multipage_key in _MULTIPAGE_CACHE:
            del _MULTIPAGE_CACHE[multipage_key]
            print(f"[DEBUG] Cleared multi-page cache for PDF: {pdf_path}")
    else:
        print(f"[DEBUG] Preserved multi-page cache for PDF: {pdf_path}")

    print(f"[DEBUG] Cleared {len(keys_to_remove)} extraction cache entries for PDF: {pdf_path}")

def clear_extraction_cache_for_section(pdf_path: str, page_number: int, section_type: str, preserve_multipage: bool = True) -> None:
    """Clear the extraction cache for a specific section of a PDF file

    Args:
        pdf_path: Path to the PDF file
        page_number: Page number (1-based)
        section_type: Section type ('header', 'items', or 'summary')
        preserve_multipage: Whether to preserve the multi-page cache (default: True)
    """
    global _EXTRACTION_CACHE
    # Find all keys that match the PDF path, page number, and section type
    keys_to_remove = []
    for key in _EXTRACTION_CACHE.keys():
        cache_key = _get_cache_key(pdf_path, page_number, None, None, section_type)
        # Check if the key starts with the hash of the PDF path, page number, and section type
        if key.startswith(cache_key[:16]):
            keys_to_remove.append(key)

    # Remove the matching keys
    for key in keys_to_remove:
        del _EXTRACTION_CACHE[key]

    print(f"[DEBUG] Cleared {len(keys_to_remove)} extraction cache entries for section {section_type} on page {page_number}")

    # Also clear the multi-page cache for this PDF, but only if not preserving it
    if not preserve_multipage:
        multipage_key = get_multipage_cache_key(pdf_path)
        if multipage_key in _MULTIPAGE_CACHE:
            del _MULTIPAGE_CACHE[multipage_key]
            print(f"[DEBUG] Cleared multi-page cache for PDF: {pdf_path}")
    else:
        print(f"[DEBUG] Preserved multi-page cache for PDF: {pdf_path} when clearing section {section_type} on page {page_number}")

def store_multipage_extraction(pdf_path: str, data: Dict) -> None:
    """Store multi-page extraction results in the cache

    Args:
        pdf_path: Path to the PDF file
        data: Dictionary containing combined data from all pages
    """
    global _MULTIPAGE_CACHE
    multipage_key = get_multipage_cache_key(pdf_path)
    _MULTIPAGE_CACHE[multipage_key] = data.copy()
    print(f"[DEBUG] Stored multi-page extraction results in cache for PDF: {pdf_path}")

def get_multipage_extraction(pdf_path: str) -> Optional[Dict]:
    """Get multi-page extraction results from the cache

    Args:
        pdf_path: Path to the PDF file

    Returns:
        Dictionary containing combined data from all pages, or None if not in cache
    """
    multipage_key = get_multipage_cache_key(pdf_path)
    if multipage_key in _MULTIPAGE_CACHE:
        print(f"[DEBUG] Using cached multi-page extraction results for PDF: {pdf_path}")
        return _MULTIPAGE_CACHE[multipage_key].copy()
    return None

# Default extraction parameters
DEFAULT_EXTRACTION_PARAMS = {
    'header': {'row_tol': 5},    # Default for header
    'items': {'row_tol': 15},    # Default for items
    'summary': {'row_tol': 10},  # Default for summary
    'split_text': True,
    'strip_text': '\n',
    'flavor': 'stream',
}

def extract_table(
    pdf_path: str,
    page_number: int = 1,
    table_area: List[float] = None,
    columns: List[float] = None,
    section_type: str = 'items',
    extraction_params: Dict = None,
    additional_params: Dict = None,
    use_cache: bool = True,
    region_index: int = 0,
    region_label: str = None
) -> pd.DataFrame:
    """
    Extract a single table from a PDF file.

    Args:
        pdf_path: Path to the PDF file
        page_number: Page number to extract from (1-based)
        table_area: Table area coordinates [x0, y0, x1, y1]
        columns: Column coordinates as a list of x-coordinates
        section_type: Section type ('header', 'items', or 'summary')
        extraction_params: Custom extraction parameters
        additional_params: Additional parameters to pass to pypdf_table_extraction
        use_cache: Whether to use the extraction cache (default: True)
        region_index: Index of the region within the section (0-based, default: 0)
        region_label: Custom label for the region (default: None)

    Returns:
        DataFrame containing the extracted table data with region labels
    """
    print(f"[DEBUG] extract_table called for {section_type} on page {page_number}")

    # Check cache first if enabled
    if use_cache:
        cache_key = _get_cache_key(pdf_path, page_number, table_area, columns, section_type, region_label)
        if cache_key in _EXTRACTION_CACHE:
            print(f"[DEBUG] Using cached extraction result for {section_type} on page {page_number}")
            return _EXTRACTION_CACHE[cache_key]

    # Check if PDF exists
    if not os.path.exists(pdf_path):
        print(f"[ERROR] PDF file does not exist: {pdf_path}")
        return pd.DataFrame()

    # Use extraction parameters exactly as provided without adding defaults
    if extraction_params is None:
        extraction_params = {
            'header': {},
            'items': {},
            'summary': {},
            'flavor': 'stream'  # Only add flavor as it's required
        }

    # Verify extraction parameters structure
    if not isinstance(extraction_params, dict):
        print(f"[WARNING] Extraction parameters are not a dictionary: {type(extraction_params)}")
        extraction_params = {
            'header': {},
            'items': {},
            'summary': {},
            'flavor': 'stream'  # Only add flavor as it's required
        }

    # Prepare parameters for pypdf_table_extraction
    # Convert table_area from list to string if it's a list
    table_area_str = None
    if table_area:
        if isinstance(table_area, list):
            table_area_str = ','.join(map(str, table_area))
        else:
            table_area_str = table_area

    # Convert columns from list to string if it's a list
    columns_str = None
    if columns:
        if isinstance(columns, list):
            columns_str = ','.join(map(str, columns))
        else:
            columns_str = columns

    # Initialize with only required parameters
    params = {
        'pages': str(page_number),
        'table_areas': [table_area_str] if table_area_str else None,
    }

    # Get section-specific parameters if available
    section_params = extraction_params.get(section_type, {})

    # Determine the flavor - first check section-specific, then global, then default to 'stream'
    flavor = section_params.get('flavor', extraction_params.get('flavor', 'stream'))

    # Only add columns parameter if flavor is 'stream'
    # This prevents the "columns cannot be used with flavor='lattice'" error
    if flavor == 'stream' and columns_str:
        params['columns'] = [columns_str]
    elif flavor == 'lattice' and columns_str:
        print(f"[WARNING] Columns parameter ignored for lattice flavor. Using lattice without columns.")

    # Set the flavor parameter
    params['flavor'] = flavor

    # Add ALL section-specific parameters directly to the params
    for key, value in section_params.items():
        if key != 'flavor':  # We've already handled flavor
            params[key] = value

    # Fall back to global parameters if section-specific ones are not available
    if 'split_text' not in params and 'split_text' in extraction_params:
        params['split_text'] = extraction_params['split_text']

    if 'strip_text' not in params and 'strip_text' in extraction_params:
        params['strip_text'] = extraction_params['strip_text']
        print(f"[DEBUG] Using strip_text: {extraction_params['strip_text']}")

    # Add any additional parameters
    if additional_params:
        params.update(additional_params)

    try:
        # Extract table using pypdf_table_extraction
        table_result = pypdf_table_extraction.read_pdf(pdf_path, **params)

        if table_result and len(table_result) > 0 and hasattr(table_result[0], "df"):
            table_df = table_result[0].df

            if table_df is not None and not table_df.empty:
                # IMMEDIATELY add page number to the DataFrame after extraction
                # This ensures the page numbers are available when generating region labels
                table_df['page_number'] = page_number
                table_df['_page_number'] = page_number

                # Add row numbers immediately as well
                # Add a temporary row number column that will be used for region labels
                table_df['_row_number'] = range(1, len(table_df) + 1)

                # Clean up the DataFrame and add region information
                table_df = clean_dataframe(table_df, section_type, region_index, region_label)

                if not table_df.empty:
                    # Store in cache if enabled
                    if use_cache:
                        cache_key = _get_cache_key(pdf_path, page_number, table_area, columns, section_type, region_label)
                        _EXTRACTION_CACHE[cache_key] = table_df.copy()
                    return table_df

        return pd.DataFrame()
    except Exception as e:
        handle_exception(
            func_name="extract_table",
            exception=e,
            context={
                "pdf_path": pdf_path,
                "page_number": page_number,
                "section_type": section_type,
                "region_label": region_label
            }
        )
        return pd.DataFrame()

def clean_dataframe(df: pd.DataFrame, region_name: str = None, region_index: int = 0, region_label: str = None) -> pd.DataFrame:
    """
    Clean a DataFrame by replacing empty strings with NaN.
    Also adds region name and row number information to the DataFrame.

    NOTE: Data cleaning operations (dropna) are commented out to preserve raw extraction results.

    Args:
        df: DataFrame to clean
        region_name: Name of the region (e.g., 'header', 'items', 'summary')
        region_index: Index of the region (0-based)
        region_label: The actual label of the region (e.g., 'H1', 'I2', 'S1')

    Returns:
        Cleaned DataFrame with region name and row number information
    """
    if df.empty:
        return df

    # Replace empty strings with NaN
    df = df.replace(r'^\s*$', pd.NA, regex=True)

    # COMMENTED OUT: Data cleaning operations to preserve raw extraction results
    # These operations were removing empty rows and columns from extracted table data
    # df = df.dropna(how="all")  # Remove rows where all values are NaN
    # df = df.dropna(axis=1, how="all")  # Remove columns where all values are NaN

    # Add region name and row number information
    if region_name:
        # Use the provided region label if available
        if region_label:
            label_prefix = region_label
        else:
            # Determine the section prefix
            if region_name == 'header':
                prefix = 'H'
            elif region_name == 'items':
                prefix = 'I'
            elif region_name == 'summary':
                prefix = 'S'
            else:
                prefix = 'X'  # Unknown region type

            # Use region_index + 1 to make it 1-based for display
            region_number = region_index + 1
            label_prefix = f"{prefix}{region_number}"

        # Add row labels with page numbers
        if '_row_number' in df.columns and '_page_number' in df.columns:
            # Use the pre-computed row numbers and page numbers
            row_labels = [f"{label_prefix}_R{int(df.iloc[i]['_row_number'])}_P{int(df.iloc[i]['_page_number'])}" for i in range(len(df))]
        elif '_row_number' in df.columns and 'page_number' in df.columns:
            # Use the pre-computed row numbers and regular page numbers
            row_labels = [f"{label_prefix}_R{int(df.iloc[i]['_row_number'])}_P{int(df.iloc[i]['page_number'])}" for i in range(len(df))]
        elif 'page_number' in df.columns:
            # Use the page_number column for each row
            row_labels = [f"{label_prefix}_R{i+1}_P{int(df.iloc[i]['page_number'])}" for i in range(len(df))]
        else:
            # Default to page 1 if no page number column is available
            row_labels = [f"{label_prefix}_R{i+1}_P1" for i in range(len(df))]

        # Add the region_label column
        df['region_label'] = row_labels

    return df

def extract_tables(
    pdf_path: str,
    page_number: int = 1,
    table_areas: List[List[float]] = None,
    columns_list: List[List[float]] = None,
    section_type: str = 'items',
    extraction_params: Dict = None,
    additional_params: Dict = None,
    use_cache: bool = True
) -> List[pd.DataFrame]:
    """
    Extract multiple tables from a PDF file.

    Args:
        pdf_path: Path to the PDF file
        page_number: Page number to extract from (1-based)
        table_areas: List of table area coordinates [[x0, y0, x1, y1], ...]
        columns_list: List of column coordinates for each table
        section_type: Section type ('header', 'items', or 'summary')
        extraction_params: Custom extraction parameters
        additional_params: Additional parameters to pass to pypdf_table_extraction
        use_cache: Whether to use the extraction cache (default: True)

    Returns:
        List of DataFrames containing the extracted table data
    """
    if not table_areas:
        return []

    # Ensure columns_list matches table_areas length
    if not columns_list:
        columns_list = [None] * len(table_areas)
    elif len(columns_list) != len(table_areas):
        columns_list = columns_list + [None] * (len(table_areas) - len(columns_list))

    processed_tables = []

    for i, (table_area, columns) in enumerate(zip(table_areas, columns_list)):
        # Create a region label based on section type and index
        titles = {'header': 'H', 'items': 'I', 'summary': 'S'}
        prefix = titles.get(section_type, section_type[0].upper() if section_type else 'X')
        region_label = f"{prefix}{i+1}"  # Add region number (1-based)

        df = extract_table(
            pdf_path=pdf_path,
            page_number=page_number,
            table_area=table_area,
            columns=columns,
            section_type=section_type,
            extraction_params=extraction_params,
            additional_params=additional_params,
            use_cache=use_cache,
            region_index=i,  # Pass the region index
            region_label=region_label  # Pass the region label
        )

        if not df.empty:
            processed_tables.append(df)

    return processed_tables


def convert_display_to_pdf_coords(display_coords, scale_x, scale_y, page_height):
    """
    Convert display coordinates to PDF coordinates with proper y-coordinate flipping.

    Args:
        display_coords: Display coordinates [x0, y0, x1, y1] where y0 is the top coordinate
        scale_x: X scale factor
        scale_y: Y scale factor
        page_height: PDF page height

    Returns:
        PDF coordinates [x0, y0, x1, y1] where y0 is the bottom coordinate
    """
    # Convert to PDF coordinates with proper y-coordinate flipping
    # In PDF coordinates, y increases from bottom to top
    return [
        display_coords[0] * scale_x,                    # x1
        page_height - (display_coords[1] * scale_y),     # y1 (flipped)
        display_coords[2] * scale_x,                    # x2
        page_height - (display_coords[3] * scale_y)      # y2 (flipped)
    ]


def convert_pdf_to_display_coords(pdf_coords, scale_x, scale_y, page_height):
    """
    Convert PDF coordinates to display coordinates with proper y-coordinate flipping.

    Args:
        pdf_coords: PDF coordinates [x0, y0, x1, y1] where y0 is the bottom coordinate
        scale_x: X scale factor
        scale_y: Y scale factor
        page_height: PDF page height

    Returns:
        Display coordinates [x0, y0, x1, y1] where y0 is the top coordinate
    """
    # Convert to display coordinates with proper y-coordinate flipping
    # In display coordinates, y increases from top to bottom
    return [
        pdf_coords[0] / scale_x,                    # x1
        (page_height - pdf_coords[1]) / scale_y,     # y1 (flipped)
        pdf_coords[2] / scale_x,                    # x2
        (page_height - pdf_coords[3]) / scale_y      # y2 (flipped)
    ]


def get_scale_factors(pdf_path: str, page_index: int = 0) -> Dict[str, float]:
    """
    Calculate scale factors for converting between display and PDF coordinates.

    Args:
        pdf_path: Path to the PDF file
        page_index: Page index (0-based)

    Returns:
        Dictionary containing scale factors
    """
    try:
        with fitz.open(pdf_path) as pdf:
            page = pdf[page_index]

            # Get actual page dimensions in points (1/72 inch)
            page_width = page.mediabox.width
            page_height = page.mediabox.height

            # Get the rendered dimensions
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            rendered_width = pix.width
            rendered_height = pix.height

            # Calculate scaling factors
            scale_x = page_width / rendered_width
            scale_y = page_height / rendered_height

            print(f"\nScaling Information:")
            print(f"  PDF Size: {page_width} x {page_height}")
            print(f"  Pixmap Scale Factor: 2")
            print(f"  Scale X: {scale_x:.4f}, Scale Y: {scale_y:.4f}")

            return {
                'scale_x': scale_x,
                'scale_y': scale_y,
                'rendered_width': rendered_width,
                'rendered_height': rendered_height,
                'page_width': page_width,
                'page_height': page_height
            }
    except Exception as e:
        handle_exception(
            func_name="get_scale_factors",
            exception=e,
            context={"pdf_path": pdf_path, "page_number": page_number}
        )
        return {
            'scale_x': 1.0,
            'scale_y': 1.0,
            'error': str(e)
        }

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
        dict: Dictionary containing extracted tables by section type
    """
    # Initialize result dictionary
    result = {
        "header": [],
        "items": [],
        "summary": [],
        "extraction_status": {
            "header": "not_processed",
            "items": "not_processed",
            "summary": "not_processed",
            "overall": "not_processed"
        }
    }

    # Use default column lines if not provided
    if column_lines is None:
        column_lines = {"header": [], "items": [], "summary": []}

    # Process each section
    for section in ['header', 'items', 'summary']:
        if section in regions and regions[section]:
            section_regions = regions[section]
            section_columns = column_lines.get(section, [])

            tables = extract_tables(
                pdf_path=pdf_path,
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

    # Set overall status
    if all(status == 'success' for status in result['extraction_status'].values() if status != 'overall'):
        result['extraction_status']['overall'] = 'success'
    elif all(status == 'failed' for status in result['extraction_status'].values() if status != 'overall'):
        result['extraction_status']['overall'] = 'failed'
    else:
        result['extraction_status']['overall'] = 'partial'

    return result

