"""
Standardized Region Types - Single Format System

This module defines the ONLY rectangle format used throughout the PDF_EXTRACTOR application.
NO backward compatibility - single format everywhere.

SINGLE FORMAT:
- Internal: {'drawing_coords': QRect, 'extraction_coords': [x0, y0, x1, y1], 'name': str}
- Database: {'draw_x': int, 'draw_y': int, 'draw_width': int, 'draw_height': int,
            'extract_x0': float, 'extract_y0': float, 'extract_x1': float, 'extract_y1': float, 'name': str}

Coordinate Systems:
- Drawing: QRect with top-left origin (y increases downward) - for UI display
- Extraction: [x0, y0, x1, y1] with bottom-left origin (y increases upward) - for pypdf_table_extraction

Key Principles:
1. Single format everywhere - NO multiple format support
2. Convert ONLY at extraction boundary for pypdf_table_extraction
3. Enforce format through type hints and validation
4. NO format detection logic anywhere
"""

from PySide6.QtCore import QRect
from typing import Dict, List, Optional, Any, TypedDict, Union
from dataclasses import dataclass

# Standard Internal Format Type Definitions
class StandardRegionItem(TypedDict):
    """Standard internal region format - used throughout the application"""
    drawing_coords: QRect  # For UI display and drawing
    extraction_coords: List[float]  # [x0, y0, x1, y1] for pypdf_table_extraction
    name: str  # REQUIRED region name

class DatabaseRegionItem(TypedDict):
    """Database storage format for regions"""
    draw_x: int
    draw_y: int
    draw_width: int
    draw_height: int
    extract_x0: float
    extract_y0: float
    extract_x1: float
    extract_y1: float
    name: str  # REQUIRED region name

# Type aliases for clarity
StandardRegionList = List[StandardRegionItem]
StandardRegionDict = Dict[str, StandardRegionList]
DatabaseRegionList = List[DatabaseRegionItem]
DatabaseRegionDict = Dict[str, DatabaseRegionList]

# Extraction format for pypdf_table_extraction
ExtractionCoordinates = List[float]  # [x0, y0, x1, y1] format

@dataclass
class RegionValidationResult:
    """Result of region validation"""
    is_valid: bool
    error_message: Optional[str] = None
    standardized_region: Optional[StandardRegionItem] = None

class StandardizedRegionFactory:
    """Factory for creating standardized regions with validation"""
    
    @staticmethod
    def create_region(drawing_coords: QRect, extraction_coords: List[float], name: str) -> StandardRegionItem:
        """Create a standardized region item

        Args:
            drawing_coords: QRect for UI display (top-left origin)
            extraction_coords: [x0, y0, x1, y1] for pypdf_table_extraction (bottom-left origin)
            name: REQUIRED region name (cannot be None or empty)

        Returns:
            StandardRegionItem in the standard format

        Raises:
            ValueError: If coordinates are invalid or name is missing/empty
        """
        if not isinstance(drawing_coords, QRect):
            raise ValueError(f"Expected QRect for drawing_coords, got {type(drawing_coords)}")

        if not StandardizedRegionFactory.validate_rect(drawing_coords):
            raise ValueError(f"Invalid drawing rectangle: {drawing_coords}")

        if not isinstance(extraction_coords, list) or len(extraction_coords) != 4:
            raise ValueError(f"Extraction coords must be [x0, y0, x1, y1], got: {extraction_coords}")

        if not all(isinstance(coord, (int, float)) for coord in extraction_coords):
            raise ValueError(f"All extraction coordinates must be numbers, got: {extraction_coords}")

        if not name or not isinstance(name, str) or not name.strip():
            raise ValueError(f"Name is required and cannot be empty, got: {repr(name)}")

        return StandardRegionItem(
            drawing_coords=drawing_coords,
            extraction_coords=extraction_coords,
            name=name.strip()
        )
    
    @staticmethod
    def validate_rect(rect: QRect) -> bool:
        """Validate that a QRect has reasonable values"""
        if not isinstance(rect, QRect):
            return False
        
        # Check for reasonable bounds
        if rect.width() <= 0 or rect.height() <= 0:
            return False
        
        if rect.x() < 0 or rect.y() < 0:
            return False
        
        # Check for extremely large values (likely errors)
        if rect.x() > 10000 or rect.y() > 10000 or rect.width() > 10000 or rect.height() > 10000:
            return False
        
        return True
    
    @staticmethod
    def validate_region(region: StandardRegionItem) -> RegionValidationResult:
        """Validate a standardized region item
        
        Args:
            region: Region item to validate
            
        Returns:
            RegionValidationResult with validation status
        """
        try:
            # Check structure
            if not isinstance(region, dict):
                return RegionValidationResult(
                    is_valid=False,
                    error_message="Region must be a dictionary"
                )
            
            # Validate drawing_coords
            if 'drawing_coords' not in region:
                return RegionValidationResult(
                    is_valid=False,
                    error_message="Region must have 'drawing_coords' key"
                )

            drawing_coords = region['drawing_coords']
            if not isinstance(drawing_coords, QRect):
                return RegionValidationResult(
                    is_valid=False,
                    error_message=f"'drawing_coords' must be QRect, got {type(drawing_coords)}"
                )

            if not StandardizedRegionFactory.validate_rect(drawing_coords):
                return RegionValidationResult(
                    is_valid=False,
                    error_message=f"Invalid drawing rectangle: {drawing_coords}"
                )

            # Validate extraction_coords
            if 'extraction_coords' not in region:
                return RegionValidationResult(
                    is_valid=False,
                    error_message="Region must have 'extraction_coords' key"
                )

            extraction_coords = region['extraction_coords']
            if not isinstance(extraction_coords, list) or len(extraction_coords) != 4:
                return RegionValidationResult(
                    is_valid=False,
                    error_message=f"'extraction_coords' must be [x0, y0, x1, y1], got: {extraction_coords}"
                )

            if not all(isinstance(coord, (int, float)) for coord in extraction_coords):
                return RegionValidationResult(
                    is_valid=False,
                    error_message=f"All extraction coordinates must be numbers, got: {extraction_coords}"
                )

            # Validate name - REQUIRED
            name = region.get('name')
            if not name or not isinstance(name, str) or not name.strip():
                return RegionValidationResult(
                    is_valid=False,
                    error_message=f"Name is required and cannot be empty, got: {repr(name)}"
                )
            
            return RegionValidationResult(
                is_valid=True,
                standardized_region=region
            )
            
        except Exception as e:
            return RegionValidationResult(
                is_valid=False,
                error_message=f"Validation error: {str(e)}"
            )
    
    @staticmethod
    def validate_regions_dict(regions: StandardRegionDict) -> Dict[str, List[RegionValidationResult]]:
        """Validate a dictionary of region lists
        
        Args:
            regions: Dictionary of region lists to validate
            
        Returns:
            Dictionary mapping region types to validation results
        """
        results = {}
        
        for region_type, region_list in regions.items():
            results[region_type] = []
            
            if not isinstance(region_list, list):
                results[region_type].append(RegionValidationResult(
                    is_valid=False,
                    error_message=f"Region list for {region_type} must be a list"
                ))
                continue
            
            for i, region in enumerate(region_list):
                validation_result = StandardizedRegionFactory.validate_region(region)
                if not validation_result.is_valid:
                    validation_result.error_message = f"Region {i}: {validation_result.error_message}"
                results[region_type].append(validation_result)
        
        return results

class CoordinateConverter:
    """Converts drawing coordinates to extraction coordinates"""

    @staticmethod
    def drawing_to_extraction_coords(drawing_rect: QRect, scale_x: float, scale_y: float, page_height: float) -> List[float]:
        """Convert drawing QRect to extraction coordinates for pypdf_table_extraction

        Args:
            drawing_rect: QRect with top-left origin (for UI display)
            scale_x: X scale factor
            scale_y: Y scale factor
            page_height: PDF page height

        Returns:
            [x0, y0, x1, y1] coordinates with bottom-left origin for extraction
        """
        # Convert QRect (top-left origin) to extraction format (bottom-left origin)
        x0 = drawing_rect.x() * scale_x
        y0 = page_height - ((drawing_rect.y() + drawing_rect.height()) * scale_y)  # Flip Y and use bottom
        x1 = (drawing_rect.x() + drawing_rect.width()) * scale_x
        y1 = page_height - (drawing_rect.y() * scale_y)  # Flip Y and use top

        return [x0, y0, x1, y1]

    @staticmethod
    def create_region_with_both_coords(drawing_rect: QRect, scale_x: float, scale_y: float, page_height: float, name: str) -> StandardRegionItem:
        """Create a region with both drawing and extraction coordinates

        Args:
            drawing_rect: QRect for UI display
            scale_x: X scale factor
            scale_y: Y scale factor
            page_height: PDF page height
            name: Region name

        Returns:
            StandardRegionItem with both coordinate systems
        """
        extraction_coords = CoordinateConverter.drawing_to_extraction_coords(
            drawing_rect, scale_x, scale_y, page_height
        )

        return StandardizedRegionFactory.create_region(drawing_rect, extraction_coords, name)

# Constants for region types
REGION_TYPES = {
    'HEADER': 'header',
    'ITEMS': 'items', 
    'SUMMARY': 'summary'
}

# Default empty regions in standard format
EMPTY_STANDARD_REGIONS: StandardRegionDict = {
    'header': [],
    'items': [],
    'summary': []
}

def create_empty_regions() -> StandardRegionDict:
    """Create an empty regions dictionary in standard format"""
    return {
        'header': [],
        'items': [],
        'summary': []
    }
