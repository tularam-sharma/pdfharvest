"""
Dual Coordinate Storage System for PDF_EXTRACTOR

This module implements dual coordinate storage to eliminate conversion overhead:
- Drawing coordinates: QRect format for UI display (top-left origin)
- Extraction coordinates: [x0, y0, x1, y1] format for pypdf_table_extraction (bottom-left origin)

Key Benefits:
1. Store both formats in database - no runtime conversion
2. Reduce complexity by eliminating coordinate transformation logic
3. Improve performance by avoiding repeated calculations
4. Maintain accuracy by storing exact coordinates for each use case
5. No backward compatibility - single standardized format

Usage:
    # Create dual coordinate region
    dual_region = DualCoordinateRegion.from_ui_input(x, y, width, height, label, scale_x, scale_y, page_height)

    # Store in database
    storage = DualCoordinateStorage()
    storage.serialize_regions(regions)

    # Retrieve for UI display
    drawing_coords = dual_region.get_drawing_coordinates()

    # Retrieve for extraction
    extraction_coords = dual_region.get_extraction_coordinates()
"""

import json
from typing import Dict, List, Optional, Any, Union, Tuple
from PySide6.QtCore import QRect
from dataclasses import dataclass, asdict
from error_handler import log_error, log_warning, log_info


@dataclass
class DualCoordinateRegion:
    """Region with both drawing and extraction coordinates stored separately"""
    # Drawing coordinates (UI display) - top-left origin
    drawing_x: int
    drawing_y: int
    drawing_width: int
    drawing_height: int
    
    # Extraction coordinates (pypdf_table_extraction) - bottom-left origin
    extraction_x0: float
    extraction_y0: float
    extraction_x1: float
    extraction_y1: float
    
    # Region metadata
    label: str
    
    def __post_init__(self):
        """Validate the region after creation"""
        if not self.label or not isinstance(self.label, str):
            raise ValueError(f"label must be non-empty string, got {repr(self.label)}")
    
    @classmethod
    def from_ui_input(cls, x: int, y: int, width: int, height: int, label: str,
                     scale_x: float = 1.0, scale_y: float = 1.0, page_height: float = 842.0) -> 'DualCoordinateRegion':
        """Create dual coordinate region from UI input
        
        Args:
            x, y, width, height: UI coordinates (top-left origin)
            label: Region label (H1, I1, S1, etc.)
            scale_x, scale_y: Scale factors for extraction coordinates
            page_height: PDF page height for extraction coordinates
            
        Returns:
            DualCoordinateRegion with both coordinate systems
        """
        # Store drawing coordinates as-is
        drawing_x, drawing_y, drawing_width, drawing_height = int(x), int(y), int(width), int(height)
        
        # Calculate extraction coordinates (bottom-left origin)
        extraction_x0 = x * scale_x
        extraction_y0 = page_height - (y + height) * scale_y  # Flip Y and use bottom
        extraction_x1 = (x + width) * scale_x
        extraction_y1 = page_height - y * scale_y  # Flip Y and use top
        
        return cls(
            drawing_x=drawing_x,
            drawing_y=drawing_y,
            drawing_width=drawing_width,
            drawing_height=drawing_height,
            extraction_x0=extraction_x0,
            extraction_y0=extraction_y0,
            extraction_x1=extraction_x1,
            extraction_y1=extraction_y1,
            label=label
        )
    
    @classmethod
    def from_qrect(cls, rect: QRect, label: str, scale_x: float = 1.0, scale_y: float = 1.0, 
                  page_height: float = 842.0) -> 'DualCoordinateRegion':
        """Create dual coordinate region from QRect
        
        Args:
            rect: QRect with UI coordinates
            label: Region label
            scale_x, scale_y: Scale factors for extraction coordinates
            page_height: PDF page height for extraction coordinates
            
        Returns:
            DualCoordinateRegion with both coordinate systems
        """
        return cls.from_ui_input(rect.x(), rect.y(), rect.width(), rect.height(), 
                               label, scale_x, scale_y, page_height)
    
    def get_drawing_coordinates(self) -> QRect:
        """Get drawing coordinates as QRect for UI display
        
        Returns:
            QRect with drawing coordinates (top-left origin)
        """
        return QRect(self.drawing_x, self.drawing_y, self.drawing_width, self.drawing_height)
    
    def get_extraction_coordinates(self) -> List[float]:
        """Get extraction coordinates for pypdf_table_extraction
        
        Returns:
            [x0, y0, x1, y1] coordinates (bottom-left origin)
        """
        return [self.extraction_x0, self.extraction_y0, self.extraction_x1, self.extraction_y1]
    
    def get_extraction_table_area_string(self) -> str:
        """Get extraction coordinates as table_area string
        
        Returns:
            Comma-separated string "x0,y0,x1,y1" for pypdf_table_extraction
        """
        return f"{self.extraction_x0},{self.extraction_y0},{self.extraction_x1},{self.extraction_y1}"
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for database storage"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DualCoordinateRegion':
        """Create from dictionary (database deserialization)"""
        return cls(**data)


@dataclass
class DualCoordinateColumnLine:
    """Column line with both drawing and extraction coordinates stored separately"""
    # Drawing coordinates (UI display) - top-left origin
    drawing_start_x: int
    drawing_start_y: int
    drawing_end_x: int
    drawing_end_y: int
    
    # Extraction coordinates (pypdf_table_extraction) - bottom-left origin
    extraction_start_x: float
    extraction_start_y: float
    extraction_end_x: float
    extraction_end_y: float
    
    # Column line metadata
    label: Optional[str] = None
    
    @classmethod
    def from_ui_input(cls, start_x: int, start_y: int, end_x: int, end_y: int, 
                     scale_x: float = 1.0, scale_y: float = 1.0, page_height: float = 842.0,
                     label: Optional[str] = None) -> 'DualCoordinateColumnLine':
        """Create dual coordinate column line from UI input
        
        Args:
            start_x, start_y, end_x, end_y: UI coordinates (top-left origin)
            scale_x, scale_y: Scale factors for extraction coordinates
            page_height: PDF page height for extraction coordinates
            label: Optional column line label
            
        Returns:
            DualCoordinateColumnLine with both coordinate systems
        """
        # Store drawing coordinates as-is
        drawing_start_x, drawing_start_y = int(start_x), int(start_y)
        drawing_end_x, drawing_end_y = int(end_x), int(end_y)
        
        # Calculate extraction coordinates (bottom-left origin)
        extraction_start_x = start_x * scale_x
        extraction_start_y = page_height - start_y * scale_y  # Flip Y
        extraction_end_x = end_x * scale_x
        extraction_end_y = page_height - end_y * scale_y  # Flip Y
        
        return cls(
            drawing_start_x=drawing_start_x,
            drawing_start_y=drawing_start_y,
            drawing_end_x=drawing_end_x,
            drawing_end_y=drawing_end_y,
            extraction_start_x=extraction_start_x,
            extraction_start_y=extraction_start_y,
            extraction_end_x=extraction_end_x,
            extraction_end_y=extraction_end_y,
            label=label
        )
    
    def get_drawing_coordinates(self) -> Tuple[int, int, int, int]:
        """Get drawing coordinates for UI display
        
        Returns:
            Tuple of (start_x, start_y, end_x, end_y) for drawing
        """
        return (self.drawing_start_x, self.drawing_start_y, self.drawing_end_x, self.drawing_end_y)
    
    def get_extraction_coordinates(self) -> List[float]:
        """Get extraction coordinates for pypdf_table_extraction
        
        Returns:
            [start_x, start_y, end_x, end_y] coordinates (bottom-left origin)
        """
        return [self.extraction_start_x, self.extraction_start_y, self.extraction_end_x, self.extraction_end_y]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for database storage"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DualCoordinateColumnLine':
        """Create from dictionary (database deserialization)"""
        return cls(**data)


# Type aliases
DualRegionDict = Dict[str, List[DualCoordinateRegion]]
DualColumnDict = Dict[str, List[DualCoordinateColumnLine]]


class DualCoordinateStorage:
    """Handles storage and retrieval of dual coordinate data"""
    
    @staticmethod
    def serialize_regions(regions: DualRegionDict) -> str:
        """Serialize dual coordinate regions for database storage
        
        Args:
            regions: Dictionary of dual coordinate regions
            
        Returns:
            JSON string for database storage
        """
        try:
            db_data = {}
            for region_type, region_list in regions.items():
                db_data[region_type] = [region.to_dict() for region in region_list]
            return json.dumps(db_data)
        except Exception as e:
            log_error(f"Error serializing dual coordinate regions: {e}")
            return json.dumps({})
    
    @staticmethod
    def deserialize_regions(json_str: str) -> DualRegionDict:
        """Deserialize dual coordinate regions from database storage
        
        Args:
            json_str: JSON string from database
            
        Returns:
            Dictionary of dual coordinate regions
        """
        try:
            if not json_str:
                return {'header': [], 'items': [], 'summary': []}
            
            data = json.loads(json_str)
            regions = {}
            
            for region_type, region_list in data.items():
                regions[region_type] = []
                for region_data in region_list:
                    try:
                        region = DualCoordinateRegion.from_dict(region_data)
                        regions[region_type].append(region)
                    except Exception as e:
                        log_warning(f"Skipping invalid dual coordinate region in {region_type}: {e}")
            
            return regions
        except Exception as e:
            log_error(f"Error deserializing dual coordinate regions: {e}")
            return {'header': [], 'items': [], 'summary': []}
    
    @staticmethod
    def serialize_column_lines(column_lines: DualColumnDict) -> str:
        """Serialize dual coordinate column lines for database storage
        
        Args:
            column_lines: Dictionary of dual coordinate column lines
            
        Returns:
            JSON string for database storage
        """
        try:
            db_data = {}
            for section_type, line_list in column_lines.items():
                db_data[section_type] = [line.to_dict() for line in line_list]
            return json.dumps(db_data)
        except Exception as e:
            log_error(f"Error serializing dual coordinate column lines: {e}")
            return json.dumps({})
    
    @staticmethod
    def deserialize_column_lines(json_str: str) -> DualColumnDict:
        """Deserialize dual coordinate column lines from database storage
        
        Args:
            json_str: JSON string from database
            
        Returns:
            Dictionary of dual coordinate column lines
        """
        try:
            if not json_str:
                return {'header': [], 'items': [], 'summary': []}
            
            data = json.loads(json_str)
            column_lines = {}
            
            for section_type, line_list in data.items():
                column_lines[section_type] = []
                for line_data in line_list:
                    try:
                        line = DualCoordinateColumnLine.from_dict(line_data)
                        column_lines[section_type].append(line)
                    except Exception as e:
                        log_warning(f"Skipping invalid dual coordinate column line in {section_type}: {e}")
            
            return column_lines
        except Exception as e:
            log_error(f"Error deserializing dual coordinate column lines: {e}")
            return {'header': [], 'items': [], 'summary': []}
