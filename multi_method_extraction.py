"""
Multi-Method PDF Extraction Module

This module provides support for multiple PDF extraction methods including:
- pypdf_table_extraction (default)
- pdftotext (from invoice2data)
- tesseract OCR (from invoice2data)
"""

import os
import tempfile
import subprocess
import pandas as pd
from typing import Dict, List, Any, Optional, Union
from error_handler import log_error, log_warning, log_info
from extraction_params_utils import (
    prepare_extraction_method_params, 
    validate_extraction_method,
    SUPPORTED_EXTRACTION_METHODS
)

# Import existing extraction functions
from pdf_extraction_utils import extract_table, extract_tables

# Check for invoice2data availability
try:
    import invoice2data
    from invoice2data import extract_data
    from invoice2data.extract.loader import read_templates
    INVOICE2DATA_AVAILABLE = True
    log_info("invoice2data is available")

    # Check for specific input modules
    try:
        from invoice2data.input.pdftotext import to_text as pdftotext_to_text
        PDFTOTEXT_INPUT_AVAILABLE = True
        log_info("invoice2data pdftotext input module is available")
    except ImportError as e:
        PDFTOTEXT_INPUT_AVAILABLE = False
        log_warning(f"invoice2data pdftotext input module not available: {e}")

    try:
        from invoice2data.input.ocrmypdf import to_text as ocrmypdf_to_text
        OCRMYPDF_INPUT_AVAILABLE = True
        log_info("invoice2data ocrmypdf input module is available")
    except ImportError as e:
        OCRMYPDF_INPUT_AVAILABLE = False
        log_warning(f"invoice2data ocrmypdf input module not available: {e}")

except ImportError as e:
    INVOICE2DATA_AVAILABLE = False
    PDFTOTEXT_INPUT_AVAILABLE = False
    OCRMYPDF_INPUT_AVAILABLE = False
    log_warning(f"invoice2data not available: {e}. pdftotext and tesseract_ocr methods will use fallback implementations.")


class MultiMethodExtractor:
    """Handles extraction using multiple methods"""
    
    def __init__(self):
        self.temp_dir = None
        
    def extract_with_method(self, 
                          pdf_path: str,
                          extraction_method: str,
                          page_number: int = 1,
                          table_areas: Optional[List[List[float]]] = None,
                          columns_list: Optional[List[List[float]]] = None,
                          section_type: str = "items",
                          extraction_params: Optional[Dict[str, Any]] = None,
                          use_cache: bool = True) -> Optional[Union[pd.DataFrame, List[pd.DataFrame]]]:
        """
        Extract data using the specified method
        
        Args:
            pdf_path: Path to PDF file
            extraction_method: Method to use for extraction
            page_number: Page number (1-based)
            table_areas: List of table area coordinates
            columns_list: List of column coordinates
            section_type: Type of section being extracted
            extraction_params: Extraction parameters
            use_cache: Whether to use caching
            
        Returns:
            Extracted data as DataFrame(s) or None if extraction failed
        """
        if not validate_extraction_method(extraction_method):
            log_error(f"Invalid extraction method: {extraction_method}")
            return None
            
        # Prepare method-specific parameters
        method_params = prepare_extraction_method_params(extraction_method, extraction_params or {})
        
        try:
            if extraction_method == "pypdf_table_extraction":
                return self._extract_with_pypdf(
                    pdf_path, page_number, table_areas, columns_list, 
                    section_type, extraction_params, use_cache
                )
            elif extraction_method == "pdftotext":
                return self._extract_with_pdftotext(
                    pdf_path, page_number, table_areas, method_params
                )
            elif extraction_method == "tesseract_ocr":
                return self._extract_with_tesseract(
                    pdf_path, page_number, table_areas, method_params
                )
            elif extraction_method == "invoice2data_full":
                return self._extract_with_invoice2data_full(
                    pdf_path, page_number, table_areas, method_params
                )
            else:
                log_error(f"Unsupported extraction method: {extraction_method}")
                return None
                
        except Exception as e:
            log_error(f"Error in {extraction_method} extraction: {str(e)}")
            return None
    
    def _extract_with_pypdf(self, 
                           pdf_path: str,
                           page_number: int,
                           table_areas: Optional[List[List[float]]],
                           columns_list: Optional[List[List[float]]],
                           section_type: str,
                           extraction_params: Dict[str, Any],
                           use_cache: bool) -> Optional[Union[pd.DataFrame, List[pd.DataFrame]]]:
        """Extract using pypdf_table_extraction method"""
        log_info(f"Extracting with pypdf_table_extraction method")
        
        if not table_areas:
            log_warning("No table areas provided for pypdf extraction")
            return None
            
        try:
            if len(table_areas) > 1:
                # Multiple tables
                return extract_tables(
                    pdf_path=pdf_path,
                    page_number=page_number,
                    table_areas=table_areas,
                    columns_list=columns_list,
                    section_type=section_type,
                    extraction_params=extraction_params,
                    use_cache=use_cache
                )
            else:
                # Single table
                return extract_table(
                    pdf_path=pdf_path,
                    page_number=page_number,
                    table_area=table_areas[0],
                    columns=columns_list[0] if columns_list else None,
                    section_type=section_type,
                    extraction_params=extraction_params,
                    use_cache=use_cache
                )
        except Exception as e:
            log_error(f"pypdf_table_extraction failed: {str(e)}")
            return None
    
    def _extract_with_pdftotext(self,
                               pdf_path: str,
                               page_number: int,
                               table_areas: Optional[List[List[float]]],
                               method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Extract using pdftotext method with fallback support"""
        log_info(f"Extracting with pdftotext method")

        # Try invoice2data input module first if available
        if INVOICE2DATA_AVAILABLE and PDFTOTEXT_INPUT_AVAILABLE:
            try:
                log_info("Attempting extraction with invoice2data pdftotext input module")
                return self._extract_with_invoice2data_pdftotext(pdf_path, page_number, table_areas, method_params)
            except Exception as e:
                log_warning(f"invoice2data pdftotext input module failed: {str(e)}, falling back to direct pdftotext")

        # Use direct pdftotext command as fallback
        log_info("Using direct pdftotext command")
        return self._fallback_pdftotext_extraction(pdf_path, page_number, method_params)

    def _extract_with_invoice2data_pdftotext(self,
                                           pdf_path: str,
                                           page_number: int,
                                           table_areas: Optional[List[List[float]]],
                                           method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Extract using invoice2data's pdftotext input module"""
        try:
            # Prepare area_details for specific page and area extraction
            area_details = None
            if table_areas and len(table_areas) > 0:
                # Use the first table area for extraction
                x1, y1, x2, y2 = table_areas[0]
                area_details = {
                    'page': page_number,
                    'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2
                }

            # Extract text using invoice2data's pdftotext input module
            text_content = pdftotext_to_text(pdf_path, area_details=area_details)

            if text_content and text_content.strip():
                return self._convert_text_to_dataframe(text_content, page_number, 'pdftotext_invoice2data')

            log_warning("invoice2data pdftotext extraction produced no data")
            return None

        except Exception as e:
            log_error(f"invoice2data pdftotext input module failed: {str(e)}")
            raise

    def _convert_text_to_dataframe(self, text_content: str, page_number: int, method_name: str) -> Optional[pd.DataFrame]:
        """Convert text content to DataFrame format"""
        try:
            lines = text_content.strip().split('\n')
            if not lines:
                return None

            # Parse text into structured data
            data = []
            for line in lines:
                if line.strip():
                    # Split by multiple spaces (common in text extraction output)
                    import re
                    parts = re.split(r'\s{2,}', line.strip())  # Split on 2+ spaces
                    parts = [part.strip() for part in parts if part.strip()]
                    if parts:
                        data.append(parts)

            if not data:
                return None

            # Create DataFrame with dynamic columns
            max_cols = max(len(row) for row in data)
            columns = [f"col_{i+1}" for i in range(max_cols)]

            # Pad rows to have same number of columns
            padded_data = []
            for row in data:
                padded_row = row + [''] * (max_cols - len(row))
                padded_data.append(padded_row)

            df = pd.DataFrame(padded_data, columns=columns)

            # Add metadata to match pypdf_table_extraction format
            df.attrs['extraction_method'] = method_name
            df.attrs['page_number'] = page_number

            log_info(f"{method_name} extracted {len(df)} rows with {len(df.columns)} columns")
            return df

        except Exception as e:
            log_error(f"Text to DataFrame conversion failed: {str(e)}")
            return None

    def _fallback_pdftotext_extraction(self, pdf_path: str, page_number: int, method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Fallback pdftotext extraction using direct command"""
        try:
            log_info("Using direct pdftotext command")

            # Check if pdftotext is available
            try:
                subprocess.run(["pdftotext", "-v"], capture_output=True, check=True)
            except (subprocess.CalledProcessError, FileNotFoundError):
                log_error("pdftotext command not found. Please install poppler-utils.")
                return None

            # Create temporary directory if needed
            if not self.temp_dir:
                self.temp_dir = tempfile.mkdtemp()

            # Use pdftotext command directly
            temp_text_file = os.path.join(self.temp_dir, f"page_{page_number}.txt")

            # Build pdftotext command
            cmd = ["pdftotext"]

            # Add parameters
            if method_params.get("layout", True):
                cmd.append("-layout")
            if method_params.get("raw", False):
                cmd.append("-raw")

            # Add page specification
            cmd.extend(["-f", str(page_number), "-l", str(page_number)])

            # Add input and output files
            cmd.extend([pdf_path, temp_text_file])

            # Run pdftotext
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)

            # Read extracted text
            if os.path.exists(temp_text_file):
                with open(temp_text_file, 'r', encoding='utf-8') as f:
                    text_content = f.read()

                if text_content.strip():
                    return self._convert_text_to_dataframe(text_content, page_number, 'pdftotext_direct')

            log_warning("Direct pdftotext extraction produced no data")
            return None

        except subprocess.CalledProcessError as e:
            log_error(f"pdftotext command failed: {e}")
            return None
        except Exception as e:
            log_error(f"Fallback pdftotext extraction failed: {str(e)}")
            return None
    
    def _extract_with_tesseract(self,
                               pdf_path: str,
                               page_number: int,
                               table_areas: Optional[List[List[float]]],
                               method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Extract using tesseract OCR method with fallback support"""
        log_info(f"Extracting with tesseract OCR method")

        # Try invoice2data OCR input module first if available
        if INVOICE2DATA_AVAILABLE and OCRMYPDF_INPUT_AVAILABLE:
            try:
                log_info("Attempting extraction with invoice2data ocrmypdf input module")
                return self._extract_with_invoice2data_ocr(pdf_path, page_number, table_areas, method_params)
            except Exception as e:
                log_warning(f"invoice2data ocrmypdf input module failed: {str(e)}, falling back to direct tesseract")

        # Use direct tesseract command as fallback
        log_info("Using direct tesseract OCR")
        return self._fallback_tesseract_extraction(pdf_path, page_number, table_areas, method_params)

    def _extract_with_invoice2data_ocr(self,
                                     pdf_path: str,
                                     page_number: int,
                                     table_areas: Optional[List[List[float]]],
                                     method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Extract using invoice2data's ocrmypdf input module"""
        try:
            # Prepare area_details for specific page and area extraction
            area_details = None
            if table_areas and len(table_areas) > 0:
                # Use the first table area for extraction
                x1, y1, x2, y2 = table_areas[0]
                area_details = {
                    'page': page_number,
                    'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2
                }

            # Prepare input reader config for OCR
            input_reader_config = {
                'language': method_params.get("lang", "eng"),
                'config': method_params.get("config", "--psm 6"),
                'timeout': method_params.get("timeout", 0)
            }

            # Extract text using invoice2data's ocrmypdf input module
            text_content = ocrmypdf_to_text(
                pdf_path,
                area_details=area_details,
                input_reader_config=input_reader_config
            )

            if text_content and text_content.strip():
                return self._convert_text_to_dataframe(text_content, page_number, 'tesseract_ocr_invoice2data')

            log_warning("invoice2data ocrmypdf extraction produced no data")
            return None

        except Exception as e:
            log_error(f"invoice2data ocrmypdf input module failed: {str(e)}")
            raise

    def _extract_with_invoice2data_full(self,
                                       pdf_path: str,
                                       page_number: int,
                                       table_areas: Optional[List[List[float]]],
                                       method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Extract using full invoice2data pipeline with template matching"""
        if not INVOICE2DATA_AVAILABLE:
            log_error("invoice2data not available for full extraction")
            return None

        try:
            log_info("Attempting extraction with full invoice2data pipeline")

            # Use invoice2data's main extraction function
            result = extract_data(pdf_path)

            if result:
                # Convert invoice2data result to DataFrame format
                data = []

                # Handle different result structures
                if isinstance(result, dict):
                    # Single result - convert to list format
                    for key, value in result.items():
                        if isinstance(value, list):
                            # Handle line items or tables
                            for item in value:
                                if isinstance(item, dict):
                                    data.append(item)
                                else:
                                    data.append({key: item})
                        else:
                            data.append({key: value})
                elif isinstance(result, list):
                    # Multiple results or line items
                    data = result

                if data:
                    # Convert to DataFrame
                    df = pd.DataFrame(data)

                    # Add metadata
                    df.attrs['extraction_method'] = 'invoice2data_full'
                    df.attrs['page_number'] = page_number

                    log_info(f"invoice2data full extraction extracted {len(df)} rows with {len(df.columns)} columns")
                    return df

            log_warning("invoice2data full extraction produced no data")
            return None

        except Exception as e:
            log_error(f"invoice2data full extraction failed: {str(e)}")
            return None

    def _fallback_tesseract_extraction(self, pdf_path: str, page_number: int,
                                     table_areas: Optional[List[List[float]]],
                                     method_params: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """Fallback tesseract extraction using direct pytesseract"""
        try:
            # Check for direct tesseract dependencies
            try:
                import pytesseract
                from PIL import Image
                import pdf2image
            except ImportError:
                log_error("pytesseract, PIL, or pdf2image not available for fallback tesseract")
                return None

            log_info("Using fallback tesseract OCR")

            # Convert PDF page to image
            images = pdf2image.convert_from_path(
                pdf_path,
                first_page=page_number,
                last_page=page_number,
                dpi=300  # High DPI for better OCR
            )

            if not images:
                log_error("Failed to convert PDF page to image")
                return None

            image = images[0]

            # If table areas are specified, crop the image
            if table_areas:
                # Use the first table area for cropping
                x1, y1, x2, y2 = table_areas[0]

                # Convert PDF coordinates to image coordinates
                img_width, img_height = image.size
                # Assuming PDF coordinates are in points (72 DPI)
                scale_x = img_width / 612  # Standard PDF width in points
                scale_y = img_height / 792  # Standard PDF height in points

                # Convert and crop
                crop_x1 = int(x1 * scale_x)
                crop_y1 = int(y1 * scale_y)
                crop_x2 = int(x2 * scale_x)
                crop_y2 = int(y2 * scale_y)

                image = image.crop((crop_x1, crop_y1, crop_x2, crop_y2))

            # Perform OCR
            ocr_config = method_params.get("config", "--psm 6")
            lang = method_params.get("lang", "eng")

            text = pytesseract.image_to_string(
                image,
                lang=lang,
                config=ocr_config
            )

            # Convert OCR text to DataFrame
            if text.strip():
                return self._convert_text_to_dataframe(text, page_number, 'tesseract_ocr_direct')

            log_warning("Direct tesseract OCR extraction produced no data")
            return None

        except Exception as e:
            log_error(f"Fallback tesseract OCR extraction failed: {str(e)}")
            return None
    
    def cleanup(self):
        """Clean up temporary files"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
                log_info("Cleaned up temporary extraction files")
            except Exception as e:
                log_warning(f"Failed to clean up temporary files: {str(e)}")


# Global extractor instance
_extractor = MultiMethodExtractor()

def extract_with_method(pdf_path: str,
                       extraction_method: str,
                       page_number: int = 1,
                       table_areas: Optional[List[List[float]]] = None,
                       columns_list: Optional[List[List[float]]] = None,
                       section_type: str = "items",
                       extraction_params: Optional[Dict[str, Any]] = None,
                       use_cache: bool = True) -> Optional[Union[pd.DataFrame, List[pd.DataFrame]]]:
    """
    Convenience function for multi-method extraction
    
    Args:
        pdf_path: Path to PDF file
        extraction_method: Method to use for extraction
        page_number: Page number (1-based)
        table_areas: List of table area coordinates
        columns_list: List of column coordinates
        section_type: Type of section being extracted
        extraction_params: Extraction parameters
        use_cache: Whether to use caching
        
    Returns:
        Extracted data as DataFrame(s) or None if extraction failed
    """
    return _extractor.extract_with_method(
        pdf_path, extraction_method, page_number, table_areas, 
        columns_list, section_type, extraction_params, use_cache
    )

def cleanup_extraction():
    """Clean up extraction resources"""
    _extractor.cleanup()
