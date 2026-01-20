"""
Column utilities for Excel data processing.

This module contains reusable functions for cleaning and normalizing
Excel column names, which often contain problematic characters.
"""
import pandas as pd


def sanitize_column_name(col):
    """
    Sanitize Excel column names by removing problematic characters.
    
    Handles common Excel column name issues:
    - NaN values
    - String "nan" (from str() conversion)
    - Newlines, tabs, carriage returns
    - Multiple consecutive spaces
    - Leading/trailing whitespace
    
    Args:
        col: Column name (any type, typically str or NaN)
        
    Returns:
        str: Sanitized column name, or empty string if invalid
        
    Examples:
        >>> sanitize_column_name("  User\\nName  ")
        'User Name'
        >>> sanitize_column_name(float('nan'))
        ''
        >>> sanitize_column_name("Email   Address")
        'Email Address'
    """
    if pd.isna(col):
        return ""
    
    s = str(col).strip()
    if s.lower() == 'nan':
        return ""
    
    # Remove newlines, tabs, carriage returns
    s = s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    
    # Normalize multiple spaces to single space
    s = ' '.join(s.split())
    
    return s
