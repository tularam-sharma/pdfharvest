"""
Clean Region Utilities - NO Backward Compatibility

This module provides utilities for the SINGLE standardized rectangle format:
- Drawing coordinates: QRect for UI display
- Extraction coordinates: [x0, y0, x1, y1] for pypdf_table_extraction  
- Region name: Required string identifier

SINGLE FORMAT:
{'drawing_coords': QRect, 'extraction_coords': [x0, y0, x1, y1], 'name': str}

NO backward compatibility - clean implementation only.
"""

from PySide6.QtCore import QRect
from typing import Dict, List, Tuple
from standardized_region_types import StandardRegionItem, StandardRegionDict
from error_handler import log_error, log_warning

def get_drawing_coordinates(region: StandardRegionItem) -> QRect:
    """Get drawing coordinates from standard region
    
    Args:
        region: StandardRegionItem
        
    Returns:
        QRect for UI display
        
    Raises:
        ValueError: If region format is invalid
    """
    if not isinstance(region, dict) or 'drawing_coords' not in region:
        raise ValueError(f"Region must have 'drawing_coords' key, got: {type(region)}")
    
    drawing_coords = region['drawing_coords']
    if not isinstance(drawing_coords, QRect):
        raise ValueError(f"'drawing_coords' must be QRect, got: {type(drawing_coords)}")
    
    return drawing_coords

def get_extraction_coordinates(region: StandardRegionItem) -> List[float]:
    """Get extraction coordinates from standard region
    
    Args:
        region: StandardRegionItem
        
    Returns:
        [x0, y0, x1, y1] coordinates for pypdf_table_extraction
        
    Raises:
        ValueError: If region format is invalid
    """
    if not isinstance(region, dict) or 'extraction_coords' not in region:
        raise ValueError(f"Region must have 'extraction_coords' key, got: {type(region)}")
    
    extraction_coords = region['extraction_coords']
    if not isinstance(extraction_coords, list) or len(extraction_coords) != 4:
        raise ValueError(f"'extraction_coords' must be [x0, y0, x1, y1], got: {extraction_coords}")
    
    return extraction_coords

def get_region_name(region: StandardRegionItem) -> str:
    """Get region name from standard region
    
    Args:
        region: StandardRegionItem
        
    Returns:
        Region name string
        
    Raises:
        ValueError: If region format is invalid or name is missing
    """
    if not isinstance(region, dict) or 'name' not in region:
        raise ValueError(f"Region must have 'name' key, got: {type(region)}")
    
    name = region['name']
    if not name or not isinstance(name, str) or not name.strip():
        raise ValueError(f"Name must be non-empty string, got: {repr(name)}")
    
    return name

def get_drawing_rect_coordinates(region: StandardRegionItem) -> Tuple[int, int, int, int]:
    """Get drawing rectangle coordinates as tuple
    
    Args:
        region: StandardRegionItem
        
    Returns:
        Tuple of (x, y, width, height) for drawing
    """
    drawing_coords = get_drawing_coordinates(region)
    return drawing_coords.x(), drawing_coords.y(), drawing_coords.width(), drawing_coords.height()

def get_extraction_table_area_string(region: StandardRegionItem) -> str:
    """Get extraction coordinates as table_area string
    
    Args:
        region: StandardRegionItem
        
    Returns:
        Comma-separated string "x0,y0,x1,y1" for pypdf_table_extraction
    """
    extraction_coords = get_extraction_coordinates(region)
    return ','.join(map(str, extraction_coords))

def create_region_item(drawing_coords: QRect, extraction_coords: List[float], name: str) -> StandardRegionItem:
    """Create a standard region item
    
    Args:
        drawing_coords: QRect for UI display
        extraction_coords: [x0, y0, x1, y1] for extraction
        name: REQUIRED region name
        
    Returns:
        StandardRegionItem in standard format
        
    Raises:
        ValueError: If any parameter is invalid
    """
    from standardized_region_types import StandardizedRegionFactory
    return StandardizedRegionFactory.create_region(drawing_coords, extraction_coords, name)

def count_regions(regions: StandardRegionDict) -> Dict[str, int]:
    """Count regions by type
    
    Args:
        regions: StandardRegionDict
        
    Returns:
        Dictionary mapping region types to counts
    """
    return {region_type: len(region_list) for region_type, region_list in regions.items()}

def get_all_regions_flat(regions: StandardRegionDict) -> List[StandardRegionItem]:
    """Get all regions as a flat list
    
    Args:
        regions: StandardRegionDict
        
    Returns:
        List of all StandardRegionItems
    """
    all_regions = []
    for region_list in regions.values():
        all_regions.extend(region_list)
    return all_regions

def filter_regions_by_name(regions: StandardRegionDict, name_pattern: str) -> StandardRegionDict:
    """Filter regions by name pattern
    
    Args:
        regions: StandardRegionDict
        name_pattern: Pattern to match (simple string contains)
        
    Returns:
        Filtered StandardRegionDict
    """
    filtered = {}
    
    for region_type, region_list in regions.items():
        filtered[region_type] = []
        
        for region in region_list:
            name = get_region_name(region)
            if name_pattern in name:
                filtered[region_type].append(region)
    
    return filtered

def get_regions_by_type(regions: StandardRegionDict, region_type: str) -> List[StandardRegionItem]:
    """Get regions of specific type
    
    Args:
        regions: StandardRegionDict
        region_type: Type to get ('header', 'items', 'summary')
        
    Returns:
        List of StandardRegionItems of specified type
    """
    return regions.get(region_type, [])

def add_region_to_dict(regions: StandardRegionDict, region_type: str, region: StandardRegionItem) -> StandardRegionDict:
    """Add region to regions dictionary
    
    Args:
        regions: StandardRegionDict
        region_type: Type to add to
        region: StandardRegionItem to add
        
    Returns:
        Updated StandardRegionDict
    """
    if region_type not in regions:
        regions[region_type] = []
    
    regions[region_type].append(region)
    return regions

def remove_region_from_dict(regions: StandardRegionDict, region_type: str, index: int) -> StandardRegionDict:
    """Remove region from regions dictionary
    
    Args:
        regions: StandardRegionDict
        region_type: Type to remove from
        index: Index to remove
        
    Returns:
        Updated StandardRegionDict
    """
    if region_type in regions and 0 <= index < len(regions[region_type]):
        regions[region_type].pop(index)
    
    return regions

def create_empty_regions() -> StandardRegionDict:
    """Create empty regions dictionary in standard format
    
    Returns:
        Empty StandardRegionDict
    """
    return {
        'header': [],
        'items': [],
        'summary': []
    }

def validate_regions_structure(regions: StandardRegionDict) -> bool:
    """Validate regions dictionary structure
    
    Args:
        regions: StandardRegionDict to validate
        
    Returns:
        True if structure is valid
    """
    try:
        if not isinstance(regions, dict):
            return False
        
        for region_type, region_list in regions.items():
            if not isinstance(region_list, list):
                return False
            
            for region in region_list:
                # This will raise ValueError if invalid
                get_drawing_coordinates(region)
                get_extraction_coordinates(region)
                get_region_name(region)
        
        return True
        
    except Exception:
        return False

def print_regions_summary(regions: StandardRegionDict, title: str = "Regions Summary") -> None:
    """Print a summary of regions for debugging
    
    Args:
        regions: StandardRegionDict to summarize
        title: Title for the summary
    """
    print(f"\n{title}")
    print("=" * len(title))
    
    total_regions = 0
    
    for region_type, region_list in regions.items():
        count = len(region_list)
        total_regions += count
        print(f"{region_type}: {count} regions")
        
        for i, region in enumerate(region_list):
            drawing_coords = get_drawing_coordinates(region)
            extraction_coords = get_extraction_coordinates(region)
            name = get_region_name(region)
            
            print(f"  {i}: '{name}'")
            print(f"     Drawing: QRect({drawing_coords.x()}, {drawing_coords.y()}, {drawing_coords.width()}, {drawing_coords.height()})")
            print(f"     Extraction: [{extraction_coords[0]}, {extraction_coords[1]}, {extraction_coords[2]}, {extraction_coords[3]}]")
    
    print(f"\nTotal: {total_regions} regions")

# Convenience functions for common operations
def get_header_regions(regions: StandardRegionDict) -> List[StandardRegionItem]:
    """Get header regions"""
    return get_regions_by_type(regions, 'header')

def get_items_regions(regions: StandardRegionDict) -> List[StandardRegionItem]:
    """Get items regions"""
    return get_regions_by_type(regions, 'items')

def get_summary_regions(regions: StandardRegionDict) -> List[StandardRegionItem]:
    """Get summary regions"""
    return get_regions_by_type(regions, 'summary')

def has_regions(regions: StandardRegionDict) -> bool:
    """Check if regions dictionary has any regions"""
    return any(len(region_list) > 0 for region_list in regions.values())

def get_first_region(regions: StandardRegionDict, region_type: str) -> StandardRegionItem:
    """Get first region of specified type
    
    Args:
        regions: StandardRegionDict
        region_type: Type to get first region from
        
    Returns:
        First StandardRegionItem of specified type
        
    Raises:
        IndexError: If no regions of specified type exist
    """
    region_list = get_regions_by_type(regions, region_type)
    if not region_list:
        raise IndexError(f"No regions of type '{region_type}' found")
    return region_list[0]
