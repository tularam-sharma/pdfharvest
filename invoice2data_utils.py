import os
import re
import json
import yaml
import pandas as pd
import numpy as np
from datetime import datetime
import shutil
import tempfile

# Check if invoice2data is available
try:
    import invoice2data
    INVOICE2DATA_AVAILABLE = True
except ImportError:
    INVOICE2DATA_AVAILABLE = False


def clean_temp_dir(directory):
    """Remove any existing files to avoid conflicts"""
    try:
        # Remove all files in the directory
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        print(f"Cleaned temporary directory: {directory}")
    except Exception as e:
        print(f"Error cleaning temporary directory: {str(e)}")


def _dataframe_to_text(data, region_type='header', region_prefix=''):
    """Convert a DataFrame or list of DataFrames to text format

    Args:
        data: DataFrame or list of DataFrames
        region_type: Type of region ('header', 'items', 'summary')
        region_prefix: Optional prefix for region labels

    Returns:
        list: List of text lines
    """
    text_lines = []

    # Handle single DataFrame
    if isinstance(data, pd.DataFrame):
        data = [data]

    # Process list of DataFrames
    for region_index, df in enumerate(data):
        if df is None or df.empty:
            continue

        # Process each row
        for row_idx, (_, row) in enumerate(df.iterrows(), 1):
            # Get region label - always use the existing label if available
            if 'region_label' in df.columns and not pd.isna(row['region_label']):
                region_label = str(row['region_label'])
            else:
                # Get section prefix
                section_prefix = ""
                if region_type == 'header':
                    section_prefix = "H"
                elif region_type == 'items':
                    section_prefix = "I"
                elif region_type == 'summary':
                    section_prefix = "S"

                # Get page number
                page_num = 1
                if '_page_number' in df.columns and not pd.isna(row['_page_number']):
                    page_num = int(row['_page_number'])
                elif 'page_number' in df.columns and not pd.isna(row['page_number']):
                    page_num = int(row['page_number'])

                # Create a region label with page number to ensure uniqueness
                # IMPORTANT: Preserve the original region number (region_index+1)
                region_label = f"{section_prefix}{region_index+1}_R{row_idx}_P{page_num}"

            # Format values, excluding metadata columns
            formatted_values = []
            columns_to_exclude = ['region_label', '_page_number', 'page_number', 'pdf_page', '_row_number']

            for col, val in row.items():
                if col in columns_to_exclude:
                    continue

                if pd.isna(val):
                    formatted_values.append("")
                elif isinstance(val, (float, np.float64, np.float32)):
                    formatted_values.append(f"{val:.2f}".rstrip('0').rstrip('.'))
                else:
                    formatted_values.append(str(val).strip())

            # Create the line without any trailing suffixes
            line = f"{region_label}|" + "|".join(formatted_values)

            # Debug: Check for unexpected |1 suffix
            if line.endswith("|1"):
                print(f"[DEBUG] Found |1 suffix in line: {line}")
                print(f"[DEBUG] Formatted values: {formatted_values}")
                print(f"[DEBUG] DataFrame columns: {list(df.columns)}")
                print(f"[DEBUG] Row data: {dict(row)}")

            text_lines.append(line)

    return text_lines


def convert_extraction_to_text(extraction_data, pdf_path=None, for_display=False):
    """Convert the extraction data to a text format that can be used by invoice2data

    Args:
        extraction_data (dict): The extracted data from pypdf_table_extraction
        pdf_path (str, optional): Path to the PDF file, used for adding metadata if not present
        for_display (bool): If True, returns clean format without pipe separators

    Returns:
        str: Text representation of the extracted data with pipe-separated values
    """
    if not extraction_data:
        return ""

    # Make a copy of the extraction data to avoid modifying the original
    extraction_data = extraction_data.copy()

    # Use standardized metadata handling
    try:
        from standardized_metadata_handler import create_standard_metadata

        # Determine template type from data structure
        # Check if this is truly multi-page by looking at page numbers in the data
        template_type = "single"
        max_page_number = 1

        for section in ['header', 'items', 'summary']:
            if section in extraction_data and extraction_data[section]:
                section_data = extraction_data[section]
                if isinstance(section_data, list):
                    for df in section_data:
                        if isinstance(df, pd.DataFrame) and 'page_number' in df.columns:
                            max_page_in_df = df['page_number'].max()
                            if max_page_in_df > max_page_number:
                                max_page_number = max_page_in_df
                elif isinstance(section_data, pd.DataFrame) and 'page_number' in section_data.columns:
                    max_page_in_df = section_data['page_number'].max()
                    if max_page_in_df > max_page_number:
                        max_page_number = max_page_in_df

        # Only set to multi if we actually have multiple pages
        if max_page_number > 1:
            template_type = "multi"

        # Create or update metadata with standardized format
        if 'metadata' not in extraction_data:
            extraction_data['metadata'] = {}

        # Update with standardized metadata
        standard_metadata = create_standard_metadata(pdf_path or "", template_type)
        extraction_data['metadata'].update(standard_metadata)

    except ImportError:
        # Fallback to original logic if standardized handler not available
        if 'metadata' not in extraction_data:
            extraction_data['metadata'] = {}

        # Add PDF path to metadata if provided
        if pdf_path and 'filename' not in extraction_data['metadata']:
            import os
            extraction_data['metadata']['filename'] = os.path.basename(pdf_path)

        # Add creation date if not present
        if 'creation_date' not in extraction_data['metadata']:
            import datetime
            extraction_data['metadata']['creation_date'] = datetime.datetime.now().isoformat()

    # Convert to text format
    text_lines = []

    # Add metadata section
    if extraction_data['metadata']:
        text_lines.append("METADATA")
        for key, value in extraction_data['metadata'].items():
            # Skip template_name field as requested
            if key != 'template_name':
                text_lines.append(f"{key}|{value}")
        text_lines.append("")

    # Use extraction_data_to_text for the rest of the sections
    if for_display:
        # For display, use clean format without pipe separators
        text_lines.extend(extraction_data_to_clean_text(extraction_data).split('\n'))
    else:
        # For invoice2data processing, use pipe-separated format
        text_lines.extend(extraction_data_to_text(extraction_data).split('\n'))

    return "\n".join(text_lines)


def extraction_data_to_text(extraction_data):
    """Convert extraction data to text format for invoice2data processing

    Args:
        extraction_data: Dictionary containing header, items, and summary data

    Returns:
        str: Text representation of the extraction data
    """
    text_lines = []

    # Add header section
    if 'header' in extraction_data and extraction_data['header']:
        header_data = extraction_data['header']
        if (isinstance(header_data, pd.DataFrame) and not header_data.empty) or \
           (isinstance(header_data, list) and any(isinstance(df, pd.DataFrame) and not df.empty for df in header_data)):
            text_lines.append("HEADER")
            text_lines.extend(_dataframe_to_text(header_data, 'header'))
            text_lines.append("")

    # Add items section
    if 'items' in extraction_data and extraction_data['items']:
        items_data = extraction_data['items']
        if (isinstance(items_data, pd.DataFrame) and not items_data.empty) or \
           (isinstance(items_data, list) and any(isinstance(df, pd.DataFrame) and not df.empty for df in items_data)):
            text_lines.append("ITEMS")
            text_lines.extend(_dataframe_to_text(items_data, 'items'))
            text_lines.append("")

    # Add summary section
    if 'summary' in extraction_data and extraction_data['summary']:
        summary_data = extraction_data['summary']
        if (isinstance(summary_data, pd.DataFrame) and not summary_data.empty) or \
           (isinstance(summary_data, list) and any(isinstance(df, pd.DataFrame) and not df.empty for df in summary_data)):
            text_lines.append("SUMMARY")
            text_lines.extend(_dataframe_to_text(summary_data, 'summary'))
            text_lines.append("")

    return "\n".join(text_lines)


def extraction_data_to_clean_text(extraction_data):
    """Convert extraction data to clean text format for display (without pipe separators)

    Args:
        extraction_data: Dictionary containing header, items, and summary data

    Returns:
        str: Clean text representation of the extraction data
    """
    text_lines = []

    # Add header section
    if 'header' in extraction_data and extraction_data['header']:
        header_data = extraction_data['header']
        if (isinstance(header_data, pd.DataFrame) and not header_data.empty) or \
           (isinstance(header_data, list) and any(isinstance(df, pd.DataFrame) and not df.empty for df in header_data)):
            text_lines.append("HEADER")
            text_lines.extend(_dataframe_to_clean_text(header_data, 'header'))
            text_lines.append("")

    # Add items section
    if 'items' in extraction_data and extraction_data['items']:
        items_data = extraction_data['items']
        if (isinstance(items_data, pd.DataFrame) and not items_data.empty) or \
           (isinstance(items_data, list) and any(isinstance(df, pd.DataFrame) and not df.empty for df in items_data)):
            text_lines.append("ITEMS")
            text_lines.extend(_dataframe_to_clean_text(items_data, 'items'))
            text_lines.append("")

    # Add summary section
    if 'summary' in extraction_data and extraction_data['summary']:
        summary_data = extraction_data['summary']
        if (isinstance(summary_data, pd.DataFrame) and not summary_data.empty) or \
           (isinstance(summary_data, list) and any(isinstance(df, pd.DataFrame) and not df.empty for df in summary_data)):
            text_lines.append("SUMMARY")
            text_lines.extend(_dataframe_to_clean_text(summary_data, 'summary'))
            text_lines.append("")

    return "\n".join(text_lines)


def _dataframe_to_clean_text(data, region_type):
    """Convert DataFrame(s) to clean text format without pipe separators

    Args:
        data: Single DataFrame or list of DataFrames
        region_type: Type of region ('header', 'items', 'summary')

    Returns:
        List of text lines in clean format
    """
    text_lines = []

    # Handle single DataFrame
    if isinstance(data, pd.DataFrame):
        data = [data]

    # Process list of DataFrames
    for region_index, df in enumerate(data):
        if df is None or df.empty:
            continue

        # Process each row
        for row_idx, (_, row) in enumerate(df.iterrows(), 1):
            # Get region label - always use the existing label if available
            if 'region_label' in df.columns and not pd.isna(row['region_label']):
                region_label = str(row['region_label'])
            else:
                # Get section prefix
                section_prefix = ""
                if region_type == 'header':
                    section_prefix = "H"
                elif region_type == 'items':
                    section_prefix = "I"
                elif region_type == 'summary':
                    section_prefix = "S"

                # Get page number
                page_num = 1
                if '_page_number' in df.columns and not pd.isna(row['_page_number']):
                    page_num = int(row['_page_number'])
                elif 'page_number' in df.columns and not pd.isna(row['page_number']):
                    page_num = int(row['page_number'])

                # Create a region label with page number to ensure uniqueness
                region_label = f"{section_prefix}{region_index+1}_R{row_idx}_P{page_num}"

            # Format values, excluding metadata columns
            formatted_values = []
            columns_to_exclude = ['region_label', '_page_number', 'page_number', 'pdf_page', '_row_number']

            for col, val in row.items():
                if col in columns_to_exclude:
                    continue

                # COMMENTED OUT: Skip empty values to preserve raw extraction results
                # if pd.isna(val):
                #     continue  # Skip empty values for clean display

                if pd.isna(val):
                    formatted_values.append("")  # Include empty values as empty strings
                elif isinstance(val, (float, np.float64, np.float32)):
                    formatted_values.append(f"{val:.2f}".rstrip('0').rstrip('.'))
                else:
                    formatted_values.append(str(val).strip())

            # Create clean line format: "Region: H1_R1_P1  Data: value1, value2, value3"
            if formatted_values:
                line = f"Region: {region_label}  Data: {', '.join(formatted_values)}"
                text_lines.append(line)

    return text_lines


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
    # Create template dictionary with only required fields
    template = {
        "issuer": issuer,
        "name": issuer,  # Required by invoice2data
        "template_name": issuer,  # Required by invoice2data
        "fields": {},
        "keywords": [],  # Will be populated below if not empty
        "options": {
            "languages": ["en"]  # Default language
        }
    }

    # Add keywords if provided
    if keywords:
        for keyword in keywords:
            if keyword and keyword not in template["keywords"]:
                template["keywords"].append(keyword)

    # Make sure issuer is included in keywords if not already present
    if issuer and issuer not in template["keywords"]:
        template["keywords"].append(issuer)

    # Add exclude_keywords if provided
    if exclude_keywords:
        template["exclude_keywords"] = exclude_keywords

    # Add options if provided
    if options_data:
        for key, value in options_data.items():
            if value:  # Only add non-empty values
                template["options"][key] = value

    # Add fields with proper type handling
    for field_name, field_data in fields_data.items():
        if isinstance(field_data, dict):
            # Complex field with regex and type
            template["fields"][field_name] = field_data
        else:
            # Simple field (just a regex pattern)
            template["fields"][field_name] = field_data

    # Add lines data if provided
    if lines_data:
        template["lines"] = lines_data

    return template


def save_invoice2data_template(template, file_path, format="yaml"):
    """Save the invoice2data template to a file

    Args:
        template (dict): The template dictionary
        file_path (str): The path to save the template
        format (str, optional): The format to save the template in (yaml or json)

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Ensure the directory exists
        os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)

        # Save in the specified format
        if format.lower() == "yaml":
            # Clean up empty values to improve YAML readability
            template_copy = template.copy()
            for key in list(template_copy.keys()):
                if isinstance(template_copy[key], list) and not template_copy[key]:
                    template_copy[key] = None
                elif isinstance(template_copy[key], dict):
                    # Clean up nested dictionaries
                    for nested_key in list(template_copy[key].keys()):
                        if template_copy[key][nested_key] == "" or template_copy[key][nested_key] == [] or template_copy[key][nested_key] == {}:
                            del template_copy[key][nested_key]
                    # Remove empty dictionaries
                    if not template_copy[key]:
                        template_copy[key] = None

            # Use safe_dump with explicit indentation for better readability
            with open(file_path, "w", encoding="utf-8") as f:
                yaml.safe_dump(
                    template_copy,
                    f,
                    default_flow_style=False,
                    allow_unicode=True,
                    sort_keys=False,  # Preserve key order
                    indent=2,         # Explicit indentation
                    width=80          # Line width
                )
        else:  # JSON
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(template, f, indent=2, ensure_ascii=False)

        return True
    except Exception as e:
        print(f"Error saving template: {str(e)}")
        return False


def analyze_invoice2data_warnings(warnings_text, template):
    """Analyze invoice2data warnings and suggest fixes

    Args:
        warnings_text (str): The warnings captured from invoice2data
        template (dict): The template that was used

    Returns:
        str: Suggestions for fixing the warnings
    """
    suggestions = []

    # Check for keyword matching issues
    if "No keyword matches found" in warnings_text:
        suggestions.append("- No keywords matched: Check that your keywords match text in the invoice")
        if template and 'keywords' in template:
            suggestions.append(f"  * Current keywords: {template['keywords']}")
            suggestions.append("  * Try adding more keywords that appear in the invoice text")
            suggestions.append("  * Keywords are case-sensitive by default, check capitalization")

    # Check for date parsing issues
    if "Failed to parse field date with parser regex" in warnings_text or "No date found" in warnings_text:
        date_field = template.get('fields', {}).get('date', {})
        if isinstance(date_field, dict):
            current_regex = date_field.get('regex', '')
            current_formats = template.get('options', {}).get('date_formats', [])

            suggestions.append("- For the 'date' field:")
            suggestions.append(f"  * Current regex: {current_regex}")
            if current_formats:
                suggestions.append(f"  * Current formats: {current_formats}")
            suggestions.append("  * Suggestions:")
            suggestions.append("    - Check if your regex pattern correctly captures the date in the invoice")
            suggestions.append("    - Make sure the date format in the invoice matches one of your format patterns")
            suggestions.append("    - Try a more general pattern like '(\\d{1,2}[.\\/-]\\d{1,2}[.\\/-]\\d{2,4})' to match common date formats")
            suggestions.append("    - Ensure your regex has a capturing group (parentheses) around the date part")
            suggestions.append("    - Add more date formats like '%d-%b-%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%B %d, %Y'")
        else:
            suggestions.append("- No date field defined in template")
            suggestions.append("  * Add a date field with a regex pattern to capture the date")
            suggestions.append("  * Example: 'date': {'regex': 'Date:\\s*(\\d{1,2}[.\\/-]\\d{1,2}[.\\/-]\\d{2,4})'}")
            suggestions.append("  * Add date_formats to options: 'options': {'date_formats': ['%d/%m/%Y', '%Y-%m-%d']}")

    # Check for amount parsing issues
    if "Failed to parse field amount with parser regex" in warnings_text or "No amount found" in warnings_text:
        amount_field = template.get('fields', {}).get('amount', {})
        if isinstance(amount_field, dict):
            current_regex = amount_field.get('regex', '')

            suggestions.append("- For the 'amount' field:")
            suggestions.append(f"  * Current regex: {current_regex}")
            suggestions.append("  * Suggestions:")
            suggestions.append("    - Check if your regex pattern correctly captures the amount in the invoice")
            suggestions.append("    - Make sure the regex handles currency symbols, thousands separators, and decimal points")
            suggestions.append("    - Try a pattern like 'Total[^\\d]+(\\d+\\.\\d{2})' or 'Amount[^\\d]+(\\d+,\\d{2})'")
            suggestions.append("    - For a more general approach: '(\\d+[\\.,]\\d{2})' to match any number with decimal point")
            suggestions.append("    - Ensure your regex has a capturing group (parentheses) around the amount part")
            suggestions.append("    - Check if the decimal separator in the invoice matches your regex (. or ,)")
        else:
            suggestions.append("- No amount field defined in template")
            suggestions.append("  * Add an amount field with a regex pattern to capture the total amount")
            suggestions.append("  * Example: 'amount': {'regex': 'Total[^\\d]+(\\d+\\.\\d{2})'}")

    # Check for invoice_number parsing issues
    if "Failed to parse field invoice_number with parser regex" in warnings_text or "No invoice number found" in warnings_text:
        invoice_field = template.get('fields', {}).get('invoice_number', {})
        if isinstance(invoice_field, dict):
            current_regex = invoice_field.get('regex', '')

            suggestions.append("- For the 'invoice_number' field:")
            suggestions.append(f"  * Current regex: {current_regex}")
            suggestions.append("  * Suggestions:")
            suggestions.append("    - Check if your regex pattern correctly captures the invoice number in the invoice")
            suggestions.append("    - Try a more general pattern like 'Invoice[^\\w]+(\\w+)' or 'Invoice\\s*#?\\s*([A-Za-z0-9\\-\\/]+)'")
            suggestions.append("    - Ensure your regex has a capturing group (parentheses) around the invoice number part")
        else:
            suggestions.append("- No invoice_number field defined in template")
            suggestions.append("  * Add an invoice_number field with a regex pattern to capture the invoice number")
            suggestions.append("  * Example: 'invoice_number': {'regex': 'Invoice[^\\w]+(\\w+)'}")

    # Check for regex syntax errors
    if "re.error" in warnings_text or "error in regular expression" in warnings_text:
        suggestions.append("- Regex syntax error detected: Check your regex patterns")
        suggestions.append("  * Common issues: unbalanced parentheses, unescaped special characters")
        suggestions.append("  * Remember to escape special characters with \\ (e.g., \\$, \\., \\(, \\))")

    # Check for YAML parsing errors
    if "YAML" in warnings_text and "error" in warnings_text:
        suggestions.append("- YAML parsing error: Check your template format")
        suggestions.append("  * Common issues: incorrect indentation, missing colons, invalid characters")

    # Check for field extraction issues
    if "Failed to extract" in warnings_text:
        suggestions.append("- Field extraction failed: Check that your field patterns match the invoice content")
        suggestions.append("  * Make sure your regex patterns have capturing groups (parentheses)")
        suggestions.append("  * Example: 'Total:\\s+(\\d+\\.\\d{2})' instead of 'Total:\\s+\\d+\\.\\d{2}'")

    # Check template structure
    if template:
        if 'fields' not in template or not template['fields']:
            suggestions.append("- Your template has no fields defined")
            suggestions.append("  * Add fields like 'amount', 'date', 'invoice_number'")
            suggestions.append("  * Example: 'fields': {'amount': {'regex': 'Total[^\\d]+(\\d+\\.\\d{2})'}}")

        if 'keywords' not in template or not template['keywords']:
            suggestions.append("- Your template has no keywords defined")
            suggestions.append("  * Add keywords that appear in the invoice text")
            suggestions.append("  * Example: 'keywords': ['Invoice', 'Company Name']")

        if 'options' not in template or 'date_formats' not in template.get('options', {}):
            suggestions.append("- No date_formats defined in options")
            suggestions.append("  * Add date_formats to options for better date extraction")
            suggestions.append("  * Example: 'options': {'date_formats': ['%d/%m/%Y', '%Y-%m-%d']}")

    # General suggestions if no specific issues found
    if not suggestions:
        suggestions.append("- Check the invoice text to see the exact format of the fields")
        suggestions.append("- Ensure your regex patterns have capturing groups (parentheses) around the parts you want to extract")
        suggestions.append("- For date fields, make sure the formats list includes the format used in the invoice")
        suggestions.append("- Try simplifying your patterns to match more general formats")
        suggestions.append("- Make sure your keywords appear in the invoice text")

    return "\n".join(suggestions)


def process_with_invoice2data(pdf_path, template_data, extracted_data=None, temp_dir=None):
    """Process a PDF file with invoice2data using the template and extracted data

    Args:
        pdf_path (str): Path to the PDF file
        template_data (dict): Template data from the database
        extracted_data (dict, optional): Extracted data from pypdf_table_extraction
        temp_dir (str, optional): Path to temporary directory

    Returns:
        dict: The extraction result from invoice2data
    """
    if not INVOICE2DATA_AVAILABLE:
        print("invoice2data module is not available. Skipping invoice2data processing.")
        return None

    try:
        # Check if we have a JSON template
        if "json_template" not in template_data or not template_data["json_template"]:
            print("No JSON template available for invoice2data processing.")
            return None

        # Get the template from the database
        template = template_data["json_template"]
        print(f"Using template from database: {json.dumps(template, indent=2, default=str)[:200]}...")

        # Create a temporary directory if not provided
        if not temp_dir:
            temp_dir = tempfile.mkdtemp(prefix="invoice2data_")
            print(f"Created temporary directory: {temp_dir}")

        # Clean up any existing files in the temp directory
        clean_temp_dir(temp_dir)

        # Prepare text content for invoice2data using our conversion method
        print("Converting extracted data to text format for invoice2data processing")
        # Use the convert_extraction_to_text function with pdf_path parameter
        text_content = convert_extraction_to_text(extracted_data, pdf_path=pdf_path)

        # Extract the base filename without extension from the PDF path
        pdf_basename = os.path.splitext(os.path.basename(pdf_path))[0]

        # Create a safe template name
        safe_template_name = re.sub(r'[^\w\-\.]', '_', pdf_basename)

        # Create temporary files for the template and text content
        temp_template_path = os.path.join(temp_dir, f"{safe_template_name}.yml")
        temp_text_path = os.path.join(temp_dir, f"{safe_template_name}.txt")

        # Save the text content to a file
        with open(temp_text_path, "w", encoding="utf-8") as f:
            f.write(text_content)
        print(f"Saved extracted text to {temp_text_path}")

        # Save the template as YAML
        try:
            with open(temp_template_path, "w", encoding="utf-8") as f:
                yaml.dump(template, f, default_flow_style=False, allow_unicode=True)
            print(f"Saved template to {temp_template_path} (YAML)")
        except Exception as e:
            # Fallback to JSON if YAML saving fails
            temp_template_path = os.path.join(temp_dir, f"{safe_template_name}.json")
            with open(temp_template_path, "w", encoding="utf-8") as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            print(f"Saved template to {temp_template_path} (JSON)")

        # Set up logging to capture warnings
        import io
        import logging
        invoice2data_logger = logging.getLogger("invoice2data")
        log_capture = io.StringIO()
        string_handler = logging.StreamHandler(log_capture)
        string_handler.setLevel(logging.WARNING)
        invoice2data_logger.addHandler(string_handler)

        # Try to use the invoice2data API directly
        try:
            from invoice2data import extract_data
            from invoice2data.extract.loader import read_templates

            # Load the template
            templates = read_templates(temp_dir)
            print(f"Loaded {len(templates)} templates from {temp_dir}")

            # Extract data
            result = extract_data(temp_text_path, templates=templates)
            print(f"Extraction result: {result}")

            # Return the result
            return result

        except Exception as api_error:
            print(f"Error using invoice2data API: {str(api_error)}")
            import traceback
            traceback.print_exc()

            # Try using the command-line approach as a fallback
            try:
                import subprocess

                # Build the command
                cmd = [
                    "invoice2data",
                    "--input-reader", "text",
                    "--exclude-built-in-templates",
                    "--template-folder", temp_dir,
                    "--output-format", "json",
                    "--output-name", os.path.join(temp_dir, f"{safe_template_name}_result"),
                    temp_text_path
                ]
                print(f"Running command: {' '.join(cmd)}")

                # Run the command
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    check=False
                )

                # Check if the command was successful
                if result.returncode == 0:
                    print(f"Command succeeded")

                    # Try to find the output file
                    result_file = os.path.join(temp_dir, f"{safe_template_name}_result.json")
                    if os.path.exists(result_file):
                        with open(result_file, "r", encoding="utf-8") as f:
                            result_data = json.load(f)
                            print(f"Loaded result from {result_file}")
                            return result_data

                    # If output file not found, try to parse from stdout
                    json_match = re.search(r'(\{.*\})', result.stdout, re.DOTALL)
                    if json_match:
                        json_str = json_match.group(1)
                        try:
                            result_data = json.loads(json_str)
                            print(f"Parsed JSON result from stdout")
                            return result_data
                        except json.JSONDecodeError:
                            print(f"Failed to parse JSON from stdout")
                else:
                    print(f"Command failed: {result.stderr}")
            except Exception as cmd_error:
                print(f"Error running command-line approach: {str(cmd_error)}")
                import traceback
                traceback.print_exc()

        # If we get here, all approaches failed
        print("All invoice2data extraction approaches failed")
        return None

    except Exception as e:
        error_message = str(e)
        print(f"Error in process_with_invoice2data: {error_message}")
        import traceback
        traceback.print_exc()
        return None











