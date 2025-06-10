"""
Extraction Parameters Utilities

This module provides standardized extraction parameter handling for consistent
PDF table extraction across split_screen_invoice_processor.py and bulk_processor.py.
"""

from typing import Dict, Any, Optional, List
from error_handler import log_error, log_warning, log_info

# Default extraction parameters for different sections
DEFAULT_SECTION_PARAMS = {
    'header': {
        'row_tol': 5,
        'edge_tol': 0.5,
        'flavor': 'stream',
        'split_text': True,
        'strip_text': '\n'
    },
    'items': {
        'row_tol': 15,
        'edge_tol': 0.5,
        'flavor': 'stream',
        'split_text': True,
        'strip_text': '\n'
    },
    'summary': {
        'row_tol': 10,
        'edge_tol': 0.5,
        'flavor': 'stream',
        'split_text': True,
        'strip_text': '\n'
    }
}

# Global default parameters
DEFAULT_GLOBAL_PARAMS = {
    'flavor': 'stream',
    'split_text': True,
    'strip_text': '\n',
    'edge_tol': 0.5,
    'parallel': True
}

# Supported extraction methods
SUPPORTED_EXTRACTION_METHODS = [
    'pypdf_table_extraction',
    'pdftotext',
    'tesseract_ocr',
    'invoice2data_full'
]

# Default parameters for different extraction methods
EXTRACTION_METHOD_DEFAULTS = {
    'pypdf_table_extraction': {
        'flavor': 'stream',
        'split_text': True,
        'strip_text': '\n',
        'edge_tol': 0.5,
        'parallel': True
    },
    'pdftotext': {
        'layout': True,
        'raw': False,
        'html': False,
        'xml': False,
        'bbox': False
    },
    'tesseract_ocr': {
        'lang': 'eng',
        'config': '--psm 6',
        'nice': 0,
        'timeout': 0
    },
    'invoice2data_full': {
        'template_folder': None,
        'exclude_built_in_templates': False,
        'input_reader': 'pdftotext',
        'output_format': 'dict'
    }
}

class ExtractionParamsHandler:
    """Utility class for standardizing extraction parameters"""
    
    @staticmethod
    def normalize_extraction_params(extraction_params: Optional[Dict[str, Any]]) -> Dict[str, Any]:
        """Normalize extraction parameters to standard format
        
        Args:
            extraction_params: Raw extraction parameters from UI or template
            
        Returns:
            Normalized extraction parameters dictionary
        """
        if not extraction_params:
            extraction_params = {}
        
        # Ensure extraction_params is a dictionary
        if not isinstance(extraction_params, dict):
            log_warning(f"Invalid extraction_params type: {type(extraction_params)}, using defaults")
            extraction_params = {}
        
        # Create normalized parameters
        normalized = {}
        
        # Add global parameters
        for key, default_value in DEFAULT_GLOBAL_PARAMS.items():
            normalized[key] = extraction_params.get(key, default_value)
        
        # Add section-specific parameters
        for section in ['header', 'items', 'summary']:
            if section not in normalized:
                normalized[section] = {}
            
            # Get section parameters from input
            section_params = extraction_params.get(section, {})
            if not isinstance(section_params, dict):
                log_warning(f"Invalid {section} params type: {type(section_params)}, using defaults")
                section_params = {}
            
            # Merge with defaults
            for key, default_value in DEFAULT_SECTION_PARAMS[section].items():
                normalized[section][key] = section_params.get(key, default_value)
            
            # Inherit global parameters if not specified in section
            for key in ['flavor', 'split_text', 'strip_text', 'edge_tol']:
                if key not in normalized[section]:
                    normalized[section][key] = normalized.get(key, DEFAULT_GLOBAL_PARAMS.get(key))
        
        # Handle custom parameters
        for key, value in extraction_params.items():
            if key.startswith('custom_param_') or key not in normalized:
                if key not in ['header', 'items', 'summary']:
                    normalized[key] = value
        
        log_info(f"Normalized extraction parameters", context={"sections": list(normalized.keys())})
        return normalized
    
    @staticmethod
    def prepare_section_params(extraction_params: Dict[str, Any], section: str, 
                             columns_list: Optional[List[float]] = None) -> Dict[str, Any]:
        """Prepare extraction parameters for a specific section
        
        Args:
            extraction_params: Normalized extraction parameters
            section: Section name ('header', 'items', 'summary')
            columns_list: Optional list of column coordinates
            
        Returns:
            Section-specific extraction parameters
        """
        # Get section-specific parameters
        section_params = extraction_params.get(section, {}).copy()
        
        # Add global parameters if not in section
        for key in ['flavor', 'split_text', 'strip_text', 'edge_tol', 'parallel']:
            if key not in section_params:
                section_params[key] = extraction_params.get(key, DEFAULT_GLOBAL_PARAMS.get(key))
        
        # Adjust flavor based on columns
        if columns_list and section_params.get('flavor') == 'lattice':
            log_warning(f"Columns defined but flavor is 'lattice', changing to 'stream' for {section}")
            section_params['flavor'] = 'stream'
        
        # Add flavor-specific parameters
        if section_params.get('flavor') == 'lattice':
            # Lattice-specific parameters
            for param_name in ['col_tol', 'min_rows']:
                if param_name in extraction_params.get(section, {}):
                    section_params[param_name] = extraction_params[section][param_name]
        
        # Add custom parameters for this section
        for key, value in extraction_params.items():
            if key.startswith(f'{section}_custom_param_'):
                # Extract custom parameter name and value
                if '_name' in key and key.replace('_name', '_value') in extraction_params:
                    param_name = value
                    param_value = extraction_params[key.replace('_name', '_value')]
                    section_params[param_name] = param_value
        
        # Add global custom parameters
        for key, value in extraction_params.items():
            if key.startswith('custom_param_') and not key.startswith(f'{section}_custom_param_'):
                if '_name' in key and key.replace('_name', '_value') in extraction_params:
                    param_name = value
                    param_value = extraction_params[key.replace('_name', '_value')]
                    section_params[param_name] = param_value
        
        log_info(f"Prepared {section} extraction parameters", 
                context={"section": section, "params": list(section_params.keys())})
        return section_params
    
    @staticmethod
    def create_additional_params(section_params: Dict[str, Any]) -> Dict[str, Any]:
        """Create additional parameters dictionary for extract_table/extract_tables
        
        Args:
            section_params: Section-specific parameters
            
        Returns:
            Additional parameters dictionary
        """
        additional_params = {}
        
        # Standard additional parameters
        for key in ['split_text', 'strip_text', 'flavor', 'parallel']:
            if key in section_params:
                additional_params[key] = section_params[key]
        
        # Add any custom parameters that aren't standard extraction parameters
        standard_keys = {'row_tol', 'edge_tol', 'col_tol', 'min_rows', 'pages'}
        for key, value in section_params.items():
            if key not in standard_keys and key not in additional_params:
                additional_params[key] = value
        
        return additional_params
    
    @staticmethod
    def create_extraction_params_dict(section_params: Dict[str, Any], 
                                    section: str, page_number: int) -> Dict[str, Any]:
        """Create extraction parameters dictionary for extract_table/extract_tables
        
        Args:
            section_params: Section-specific parameters
            section: Section name
            page_number: Page number (1-based)
            
        Returns:
            Extraction parameters dictionary
        """
        extraction_params = {
            'pages': str(page_number),
            section: {
                'row_tol': section_params.get('row_tol', DEFAULT_SECTION_PARAMS[section]['row_tol'])
            }
        }
        
        # Add standard parameters
        for key in ['edge_tol']:
            if key in section_params:
                extraction_params[key] = section_params[key]
        
        # Add flavor-specific parameters
        if section_params.get('flavor') == 'lattice':
            for key in ['col_tol', 'min_rows']:
                if key in section_params:
                    extraction_params[key] = section_params[key]
        
        return extraction_params
    
    @staticmethod
    def validate_extraction_params(extraction_params: Dict[str, Any]) -> bool:
        """Validate extraction parameters structure
        
        Args:
            extraction_params: Extraction parameters to validate
            
        Returns:
            True if valid, False otherwise
        """
        if not isinstance(extraction_params, dict):
            log_error("Extraction parameters must be a dictionary")
            return False
        
        # Check section parameters
        for section in ['header', 'items', 'summary']:
            if section in extraction_params:
                section_params = extraction_params[section]
                if not isinstance(section_params, dict):
                    log_error(f"Section '{section}' parameters must be a dictionary")
                    return False
                
                # Check required parameters
                if 'row_tol' in section_params:
                    try:
                        float(section_params['row_tol'])
                    except (ValueError, TypeError):
                        log_error(f"Invalid row_tol value for {section}: {section_params['row_tol']}")
                        return False
        
        # Check flavor parameter
        flavor = extraction_params.get('flavor', 'stream')
        if flavor not in ['stream', 'lattice']:
            log_error(f"Invalid flavor: {flavor}. Must be 'stream' or 'lattice'")
            return False
        
        log_info("Extraction parameters validation passed")
        return True

    @staticmethod
    def prepare_extraction_method_params(extraction_method: str,
                                       extraction_params: Dict[str, Any]) -> Dict[str, Any]:
        """Prepare parameters for specific extraction method

        Args:
            extraction_method: The extraction method to use
            extraction_params: Base extraction parameters

        Returns:
            Method-specific extraction parameters
        """
        if extraction_method not in SUPPORTED_EXTRACTION_METHODS:
            log_warning(f"Unsupported extraction method: {extraction_method}, using pypdf_table_extraction")
            extraction_method = 'pypdf_table_extraction'

        # Get method-specific defaults
        method_defaults = EXTRACTION_METHOD_DEFAULTS.get(extraction_method, {})

        # Merge with provided parameters
        method_params = method_defaults.copy()

        if extraction_method == 'pypdf_table_extraction':
            # For pypdf_table_extraction, use the standard parameter handling
            method_params.update(extraction_params)
        elif extraction_method == 'pdftotext':
            # For pdftotext, merge specific parameters
            pdftotext_params = extraction_params.get('pdftotext', {})
            method_params.update(pdftotext_params)
        elif extraction_method == 'tesseract_ocr':
            # For tesseract OCR, merge specific parameters
            tesseract_params = extraction_params.get('tesseract_ocr', {})
            method_params.update(tesseract_params)
        elif extraction_method == 'invoice2data_full':
            # For full invoice2data pipeline, merge specific parameters
            invoice2data_params = extraction_params.get('invoice2data_full', {})
            method_params.update(invoice2data_params)

        log_info(f"Prepared {extraction_method} extraction parameters",
                context={"method": extraction_method, "params": list(method_params.keys())})
        return method_params

    @staticmethod
    def validate_extraction_method(extraction_method: str) -> bool:
        """Validate extraction method

        Args:
            extraction_method: The extraction method to validate

        Returns:
            True if valid, False otherwise
        """
        if extraction_method not in SUPPORTED_EXTRACTION_METHODS:
            log_error(f"Unsupported extraction method: {extraction_method}. "
                     f"Supported methods: {', '.join(SUPPORTED_EXTRACTION_METHODS)}")
            return False

        log_info(f"Extraction method validation passed: {extraction_method}")
        return True

# Convenience functions
def normalize_extraction_params(extraction_params: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Convenience function for normalizing extraction parameters"""
    return ExtractionParamsHandler.normalize_extraction_params(extraction_params)

def prepare_section_params(extraction_params: Dict[str, Any], section: str,
                         columns_list: Optional[List[float]] = None) -> Dict[str, Any]:
    """Convenience function for preparing section parameters"""
    return ExtractionParamsHandler.prepare_section_params(extraction_params, section, columns_list)

def prepare_extraction_method_params(extraction_method: str,
                                   extraction_params: Dict[str, Any]) -> Dict[str, Any]:
    """Convenience function for preparing extraction method parameters"""
    return ExtractionParamsHandler.prepare_extraction_method_params(extraction_method, extraction_params)

def validate_extraction_method(extraction_method: str) -> bool:
    """Convenience function for validating extraction method"""
    return ExtractionParamsHandler.validate_extraction_method(extraction_method)

def create_standardized_extraction_call(pdf_path: str, page_number: int, table_areas: List[List[float]], 
                                      columns_list: Optional[List[List[float]]], section: str,
                                      extraction_params: Dict[str, Any], use_cache: bool = True) -> Dict[str, Any]:
    """Create standardized parameters for extract_table/extract_tables calls
    
    Args:
        pdf_path: Path to PDF file
        page_number: Page number (1-based)
        table_areas: List of table area coordinates
        columns_list: List of column coordinates for each table
        section: Section name ('header', 'items', 'summary')
        extraction_params: Raw extraction parameters
        use_cache: Whether to use caching
        
    Returns:
        Dictionary with standardized call parameters
    """
    # Normalize parameters
    normalized_params = normalize_extraction_params(extraction_params)
    
    # Prepare section-specific parameters
    section_params = prepare_section_params(normalized_params, section, 
                                          columns_list[0] if columns_list else None)
    
    # Create extraction parameters dictionary
    extraction_params_dict = ExtractionParamsHandler.create_extraction_params_dict(
        section_params, section, page_number)
    
    # Create additional parameters
    additional_params = ExtractionParamsHandler.create_additional_params(section_params)
    
    return {
        'pdf_path': pdf_path,
        'page_number': page_number,
        'table_areas': table_areas,
        'columns_list': columns_list,
        'section_type': section,
        'extraction_params': extraction_params_dict,
        'additional_params': additional_params,
        'use_cache': use_cache
    }
