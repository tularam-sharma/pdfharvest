"""
Standardized Coordinate System - Convert Once, Use Everywhere

This module implements a single, standardized coordinate format used throughout
the PDF_EXTRACTOR application. NO multiple format support.

SINGLE FORMAT EVERYWHERE:
- Internal: {'rect': QRect, 'label': str, 'extraction_coords': [x0, y0, x1, y1]}
- Database: {'x': int, 'y': int, 'width': int, 'height': int, 'label': str}
- UI Display: QRect (top-left origin, y increases downward)
- Extraction: [x0, y0, x1, y1] (bottom-left origin, y increases upward)

Key Principles:
1. Convert ONCE at system boundaries (UI input, database I/O)
2. Store BOTH drawing and extraction coordinates together
3. NO format detection or conversion logic in business code
4. Single format enforced through type validation
"""

import json
from typing import Dict, List, Optional, Any, Union, Tuple
from PySide6.QtCore import QRect
from dataclasses import dataclass
from error_handler import log_error, log_warning, log_info

# Type definitions for the standardized system
@dataclass
class StandardRegion:
    """Single standardized region format with dual coordinates"""
    rect: QRect  # Drawing coordinates (UI display)
    label: str   # Required label (H1, I1, S1, etc.)
    extraction_coords: List[float]  # [x0, y0, x1, y1] for pypdf_table_extraction
    
    def __post_init__(self):
        """Validate the region after creation"""
        if not isinstance(self.rect, QRect):
            raise ValueError(f"rect must be QRect, got {type(self.rect)}")
        if not self.label or not isinstance(self.label, str):
            raise ValueError(f"label must be non-empty string, got {repr(self.label)}")
        if not isinstance(self.extraction_coords, list) or len(self.extraction_coords) != 4:
            raise ValueError(f"extraction_coords must be list of 4 floats, got {self.extraction_coords}")

# Type aliases
StandardRegionDict = Dict[str, List[StandardRegion]]
DatabaseRegionItem = Dict[str, Union[int, str]]

class CoordinateConverter:
    """Handles coordinate conversion between UI and extraction formats"""
    
    @staticmethod
    def ui_to_extraction_coords(rect: QRect, scale_x: float, scale_y: float, page_height: float) -> List[float]:
        """Convert UI coordinates (QRect) to extraction coordinates
        
        Args:
            rect: QRect with UI coordinates (top-left origin)
            scale_x: X scale factor from UI to PDF
            scale_y: Y scale factor from UI to PDF  
            page_height: PDF page height in points
            
        Returns:
            [x0, y0, x1, y1] coordinates for pypdf_table_extraction (bottom-left origin)
        """
        # Convert QRect to PDF coordinates
        x0 = rect.x() * scale_x
        y0 = page_height - (rect.y() + rect.height()) * scale_y  # Flip Y and use bottom
        x1 = (rect.x() + rect.width()) * scale_x
        y1 = page_height - rect.y() * scale_y  # Flip Y and use top
        
        return [x0, y0, x1, y1]
    
    @staticmethod
    def extraction_to_ui_coords(coords: List[float], scale_x: float, scale_y: float, page_height: float) -> QRect:
        """Convert extraction coordinates to UI coordinates (QRect)
        
        Args:
            coords: [x0, y0, x1, y1] extraction coordinates (bottom-left origin)
            scale_x: X scale factor from PDF to UI
            scale_y: Y scale factor from PDF to UI
            page_height: PDF page height in points
            
        Returns:
            QRect with UI coordinates (top-left origin)
        """
        x0, y0, x1, y1 = coords
        
        # Convert to UI coordinates
        ui_x = int(x0 / scale_x)
        ui_y = int((page_height - y1) / scale_y)  # Flip Y and use top
        ui_width = int((x1 - x0) / scale_x)
        ui_height = int((y1 - y0) / scale_y)
        
        return QRect(ui_x, ui_y, ui_width, ui_height)

class StandardRegionFactory:
    """Factory for creating standardized regions"""
    
    @staticmethod
    def create_region(rect: QRect, label: str, scale_x: float = 1.0, scale_y: float = 1.0, page_height: float = 842.0) -> StandardRegion:
        """Create a standardized region with both coordinate systems
        
        Args:
            rect: QRect with UI coordinates
            label: Region label (required)
            scale_x: X scale factor for extraction coordinates
            scale_y: Y scale factor for extraction coordinates
            page_height: PDF page height for extraction coordinates
            
        Returns:
            StandardRegion with both coordinate systems
        """
        if not isinstance(rect, QRect):
            raise ValueError(f"rect must be QRect, got {type(rect)}")
        if not label or not isinstance(label, str):
            raise ValueError(f"label must be non-empty string, got {repr(label)}")
        
        # Calculate extraction coordinates
        extraction_coords = CoordinateConverter.ui_to_extraction_coords(rect, scale_x, scale_y, page_height)
        
        return StandardRegion(
            rect=rect,
            label=label,
            extraction_coords=extraction_coords
        )
    
    @staticmethod
    def from_ui_input(x: int, y: int, width: int, height: int, label: str, 
                     scale_x: float = 1.0, scale_y: float = 1.0, page_height: float = 842.0) -> StandardRegion:
        """Create region from UI input coordinates
        
        Args:
            x, y, width, height: UI coordinates
            label: Region label (required)
            scale_x, scale_y: Scale factors for extraction coordinates
            page_height: PDF page height for extraction coordinates
            
        Returns:
            StandardRegion with both coordinate systems
        """
        rect = QRect(int(x), int(y), int(width), int(height))
        return StandardRegionFactory.create_region(rect, label, scale_x, scale_y, page_height)
    
    @staticmethod
    def from_database(db_item: DatabaseRegionItem, scale_x: float = 1.0, scale_y: float = 1.0, page_height: float = 842.0) -> StandardRegion:
        """Create region from database format
        
        Args:
            db_item: Database item with x, y, width, height, label
            scale_x, scale_y: Scale factors for extraction coordinates
            page_height: PDF page height for extraction coordinates
            
        Returns:
            StandardRegion with both coordinate systems
        """
        rect = QRect(
            int(db_item['x']),
            int(db_item['y']),
            int(db_item['width']),
            int(db_item['height'])
        )
        label = db_item.get('label', '')
        if not label:
            raise ValueError("Database region must have a label")
        
        return StandardRegionFactory.create_region(rect, label, scale_x, scale_y, page_height)

class DatabaseConverter:
    """Handles conversion between standard format and database storage"""
    
    @staticmethod
    def to_database_format(region: StandardRegion) -> DatabaseRegionItem:
        """Convert standard region to database format

        Args:
            region: StandardRegion to convert

        Returns:
            Dictionary for database storage
        """
        if not isinstance(region, StandardRegion):
            raise ValueError(f"Expected StandardRegion, got {type(region)}")

        return {
            'x': region.rect.x(),
            'y': region.rect.y(),
            'width': region.rect.width(),
            'height': region.rect.height(),
            'label': region.label
        }
    
    @staticmethod
    def serialize_regions(regions: StandardRegionDict) -> str:
        """Serialize regions for database storage
        
        Args:
            regions: Dictionary of standard regions
            
        Returns:
            JSON string for database storage
        """
        try:
            db_data = {}
            for region_type, region_list in regions.items():
                db_data[region_type] = [
                    DatabaseConverter.to_database_format(region)
                    for region in region_list
                ]
            return json.dumps(db_data)
        except Exception as e:
            log_error(f"Error serializing regions for database: {e}")
            return json.dumps({})
    
    @staticmethod
    def deserialize_regions(json_str: str, scale_x: float = 1.0, scale_y: float = 1.0, page_height: float = 842.0) -> StandardRegionDict:
        """Deserialize regions from database storage
        
        Args:
            json_str: JSON string from database
            scale_x, scale_y: Scale factors for extraction coordinates
            page_height: PDF page height for extraction coordinates
            
        Returns:
            Dictionary of standard regions
        """
        try:
            if not json_str:
                return {'header': [], 'items': [], 'summary': []}
            
            data = json.loads(json_str)
            regions = {}
            
            for region_type, region_list in data.items():
                regions[region_type] = []
                for db_item in region_list:
                    try:
                        region = StandardRegionFactory.from_database(db_item, scale_x, scale_y, page_height)
                        regions[region_type].append(region)
                    except Exception as e:
                        log_warning(f"Skipping invalid region in {region_type}: {e}")
            
            return regions
        except Exception as e:
            log_error(f"Error deserializing regions from database: {e}")
            return {'header': [], 'items': [], 'summary': []}

class ExtractionConverter:
    """Handles conversion to extraction format for pypdf_table_extraction"""
    
    @staticmethod
    def get_extraction_coordinates(region: StandardRegion) -> List[float]:
        """Get extraction coordinates from standard region
        
        Args:
            region: StandardRegion
            
        Returns:
            [x0, y0, x1, y1] coordinates for pypdf_table_extraction
        """
        return region.extraction_coords.copy()
    
    @staticmethod
    def get_table_area_string(region: StandardRegion) -> str:
        """Get table_area string for pypdf_table_extraction
        
        Args:
            region: StandardRegion
            
        Returns:
            Comma-separated string "x0,y0,x1,y1"
        """
        return ','.join(map(str, region.extraction_coords))
    
    @staticmethod
    def regions_to_extraction_format(regions: StandardRegionDict) -> Dict[str, List[List[float]]]:
        """Convert all regions to extraction format
        
        Args:
            regions: Dictionary of standard regions
            
        Returns:
            Dictionary mapping region types to extraction coordinates
        """
        extraction_regions = {}
        for region_type, region_list in regions.items():
            extraction_regions[region_type] = [
                region.extraction_coords.copy()
                for region in region_list
            ]
        return extraction_regions

class RegionValidator:
    """Validates standardized regions"""
    
    @staticmethod
    def validate_region(region: StandardRegion) -> Tuple[bool, str]:
        """Validate a standard region
        
        Args:
            region: StandardRegion to validate
            
        Returns:
            (is_valid, error_message)
        """
        try:
            # Check rect validity
            if not isinstance(region.rect, QRect):
                return False, f"Invalid rect type: {type(region.rect)}"
            
            if region.rect.width() <= 0 or region.rect.height() <= 0:
                return False, f"Invalid rect dimensions: {region.rect.width()}x{region.rect.height()}"
            
            # Check label validity
            if not region.label or not isinstance(region.label, str):
                return False, f"Invalid label: {repr(region.label)}"
            
            # Check extraction coordinates
            if not isinstance(region.extraction_coords, list) or len(region.extraction_coords) != 4:
                return False, f"Invalid extraction_coords: {region.extraction_coords}"
            
            return True, ""
        except Exception as e:
            return False, f"Validation error: {e}"
    
    @staticmethod
    def validate_regions_dict(regions: StandardRegionDict) -> Tuple[bool, List[str]]:
        """Validate a regions dictionary
        
        Args:
            regions: Dictionary of standard regions
            
        Returns:
            (is_valid, list_of_errors)
        """
        errors = []
        
        # Check required keys
        required_keys = {'header', 'items', 'summary'}
        if not all(key in regions for key in required_keys):
            missing = required_keys - set(regions.keys())
            errors.append(f"Missing required region types: {missing}")
        
        # Validate each region
        for region_type, region_list in regions.items():
            if not isinstance(region_list, list):
                errors.append(f"Region type '{region_type}' must be a list, got {type(region_list)}")
                continue
            
            for i, region in enumerate(region_list):
                if not isinstance(region, StandardRegion):
                    errors.append(f"Region {region_type}[{i}] must be StandardRegion, got {type(region)}")
                    continue
                
                is_valid, error_msg = RegionValidator.validate_region(region)
                if not is_valid:
                    errors.append(f"Region {region_type}[{i}]: {error_msg}")
        
        return len(errors) == 0, errors
