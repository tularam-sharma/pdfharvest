"""
Region Utilities - DEPRECATED

This module is deprecated. All coordinate handling now uses StandardRegion objects.
Use standardized_coordinates.py and coordinate_boundary_converters.py instead.

This file is kept only for essential utilities that don't involve coordinate conversion.
"""

from PySide6.QtCore import QRect
from typing import Optional
from error_handler import log_error, log_warning

def validate_rect(rect: QRect) -> bool:
    """
    Validate that a QRect has reasonable values.

    Args:
        rect: QRect to validate

    Returns:
        True if rect is valid, False otherwise
    """
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
