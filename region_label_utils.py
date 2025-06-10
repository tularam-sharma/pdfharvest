"""
Region Label Utilities

This module provides standardized region label handling for consistent
labeling across split_screen_invoice_processor.py and bulk_processor.py.
"""

import pandas as pd
import re
from typing import Dict, List, Optional, Tuple, Any
from error_handler import log_error, log_warning, log_info

class RegionLabelHandler:
    """Utility class for standardizing region labels"""
    
    @staticmethod
    def create_region_label(section: str, region_index: int, row_index: int = 1, page_number: int = 1) -> str:
        """Create a standardized region label
        
        Args:
            section: Section type ('header', 'items', 'summary')
            region_index: Region index (0-based, will be converted to 1-based)
            row_index: Row index within the region (1-based)
            page_number: Page number (1-based)
            
        Returns:
            Standardized region label (e.g., 'H1_R1_P1')
        """
        # Get section prefix
        section_prefix = {
            'header': 'H',
            'items': 'I', 
            'summary': 'S'
        }.get(section.lower(), 'X')
        
        # Convert region_index to 1-based
        region_number = region_index + 1
        
        return f"{section_prefix}{region_number}_R{row_index}_P{page_number}"
    
    @staticmethod
    def parse_region_label(region_label: str) -> Tuple[Optional[str], Optional[int], Optional[int], Optional[int]]:
        """Parse a region label into its components
        
        Args:
            region_label: Region label to parse (e.g., 'H1_R1_P1')
            
        Returns:
            Tuple of (section, region_number, row_number, page_number) or (None, None, None, None) if invalid
        """
        try:
            # Pattern: {section_prefix}{region_number}_R{row_number}_P{page_number}
            match = re.match(r'^([HIS])(\d+)_R(\d+)_P(\d+)$', region_label)
            if match:
                section_prefix, region_num, row_num, page_num = match.groups()
                
                # Convert prefix to section name
                section = {
                    'H': 'header',
                    'I': 'items',
                    'S': 'summary'
                }.get(section_prefix)
                
                return section, int(region_num), int(row_num), int(page_num)
            else:
                log_warning(f"Invalid region label format: {region_label}")
                return None, None, None, None
                
        except Exception as e:
            log_error(f"Error parsing region label '{region_label}': {e}")
            return None, None, None, None
    
    @staticmethod
    def standardize_dataframe_labels(df: pd.DataFrame, section: str, region_index: int, page_number: int = 1) -> pd.DataFrame:
        """Standardize region labels in a DataFrame
        
        Args:
            df: DataFrame to standardize
            section: Section type ('header', 'items', 'summary')
            region_index: Region index (0-based)
            page_number: Page number (1-based)
            
        Returns:
            DataFrame with standardized region labels
        """
        if df is None or df.empty:
            return df
        
        # Make a copy to avoid modifying the original
        df_copy = df.copy()
        
        # Add page_number column if not present
        if 'page_number' not in df_copy.columns:
            df_copy['page_number'] = page_number
        
        # Create standardized region labels
        region_labels = []
        for row_idx in range(len(df_copy)):
            label = RegionLabelHandler.create_region_label(
                section=section,
                region_index=region_index,
                row_index=row_idx + 1,  # 1-based row index
                page_number=page_number
            )
            region_labels.append(label)
        
        df_copy['region_label'] = region_labels
        
        log_info(f"Standardized {len(region_labels)} region labels for {section} region {region_index + 1}")
        return df_copy
    
    @staticmethod
    def extract_clean_data_from_text_format(text_line: str) -> Tuple[Optional[str], List[str]]:
        """Extract clean data from invoice2data text format
        
        Args:
            text_line: Text line in format 'H1_R1_P1|data1|data2|...'
            
        Returns:
            Tuple of (region_label, [data_values])
        """
        try:
            if '|' not in text_line:
                return None, []
            
            parts = text_line.split('|')
            if len(parts) < 2:
                return None, []
            
            region_label = parts[0]
            data_values = [part.strip() for part in parts[1:] if part.strip()]
            
            return region_label, data_values
            
        except Exception as e:
            log_error(f"Error extracting data from text line '{text_line}': {e}")
            return None, []
    
    @staticmethod
    def convert_text_format_to_dataframe(text_lines: List[str], section: str) -> Optional[pd.DataFrame]:
        """Convert invoice2data text format back to clean DataFrame
        
        Args:
            text_lines: List of text lines in format 'H1_R1_P1|data1|data2|...'
            section: Section type for validation
            
        Returns:
            Clean DataFrame with proper region labels
        """
        try:
            if not text_lines:
                return None
            
            # Extract data from text lines
            rows_data = []
            region_labels = []
            
            for line in text_lines:
                region_label, data_values = RegionLabelHandler.extract_clean_data_from_text_format(line)
                if region_label and data_values:
                    # Validate that this is the correct section
                    parsed_section, _, _, _ = RegionLabelHandler.parse_region_label(region_label)
                    if parsed_section == section:
                        region_labels.append(region_label)
                        rows_data.append(data_values)
            
            if not rows_data:
                return None
            
            # Determine column count (use the maximum number of columns)
            max_cols = max(len(row) for row in rows_data) if rows_data else 0
            
            # Create DataFrame with generic column names
            columns = [f'col_{i+1}' for i in range(max_cols)]
            
            # Pad rows to have the same number of columns
            padded_rows = []
            for row in rows_data:
                padded_row = row + [''] * (max_cols - len(row))
                padded_rows.append(padded_row)
            
            # Create DataFrame
            df = pd.DataFrame(padded_rows, columns=columns)
            df['region_label'] = region_labels
            
            # Extract page numbers from region labels
            page_numbers = []
            for label in region_labels:
                _, _, _, page_num = RegionLabelHandler.parse_region_label(label)
                page_numbers.append(page_num if page_num else 1)
            
            df['page_number'] = page_numbers
            
            log_info(f"Converted {len(text_lines)} text lines to DataFrame with {len(df)} rows for {section}")
            return df
            
        except Exception as e:
            log_error(f"Error converting text format to DataFrame for {section}: {e}")
            return None
    
    @staticmethod
    def get_display_label(region_label: str) -> str:
        """Get a clean display label from a region label
        
        Args:
            region_label: Full region label (e.g., 'H1_R1_P1')
            
        Returns:
            Clean display label (e.g., 'H1')
        """
        try:
            section, region_num, _, _ = RegionLabelHandler.parse_region_label(region_label)
            if section and region_num:
                section_prefix = {
                    'header': 'H',
                    'items': 'I',
                    'summary': 'S'
                }.get(section, 'X')
                return f"{section_prefix}{region_num}"
            else:
                return region_label
        except:
            return region_label
    
    @staticmethod
    def validate_region_labels_consistency(data: Dict[str, Any]) -> bool:
        """Validate that region labels are consistent across sections
        
        Args:
            data: Dictionary containing extraction data
            
        Returns:
            True if labels are consistent, False otherwise
        """
        try:
            all_labels = []
            
            # Collect all region labels
            for section in ['header', 'items', 'summary']:
                if section in data and data[section]:
                    section_data = data[section]
                    
                    if isinstance(section_data, pd.DataFrame) and 'region_label' in section_data.columns:
                        all_labels.extend(section_data['region_label'].tolist())
                    elif isinstance(section_data, list):
                        for df in section_data:
                            if isinstance(df, pd.DataFrame) and 'region_label' in df.columns:
                                all_labels.extend(df['region_label'].tolist())
            
            # Validate each label
            valid_count = 0
            for label in all_labels:
                section, region_num, row_num, page_num = RegionLabelHandler.parse_region_label(label)
                if section and region_num and row_num and page_num:
                    valid_count += 1
            
            consistency_ratio = valid_count / len(all_labels) if all_labels else 1.0
            is_consistent = consistency_ratio >= 0.95  # Allow 5% tolerance
            
            log_info(f"Region label consistency: {valid_count}/{len(all_labels)} valid ({consistency_ratio:.1%})")
            
            if not is_consistent:
                log_warning(f"Region labels are not consistent: {consistency_ratio:.1%} valid")
            
            return is_consistent
            
        except Exception as e:
            log_error(f"Error validating region label consistency: {e}")
            return False

# Convenience functions
def create_region_label(section: str, region_index: int, row_index: int = 1, page_number: int = 1) -> str:
    """Convenience function for creating region labels"""
    return RegionLabelHandler.create_region_label(section, region_index, row_index, page_number)

def standardize_dataframe_labels(df: pd.DataFrame, section: str, region_index: int, page_number: int = 1) -> pd.DataFrame:
    """Convenience function for standardizing DataFrame labels"""
    return RegionLabelHandler.standardize_dataframe_labels(df, section, region_index, page_number)

def get_display_label(region_label: str) -> str:
    """Convenience function for getting display labels"""
    return RegionLabelHandler.get_display_label(region_label)
