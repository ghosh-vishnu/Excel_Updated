#!/usr/bin/env python3
"""
Test script to verify that the report coverage table extraction
now uses inline CSS instead of internal CSS.
"""

import sys
import os
sys.path.append('/Users/ravishkumar/Desktop/word-conveter/Excel_Updated/backend')

from converter.utils.extractor import extract_report_coverage_table_with_style

def test_inline_css_conversion():
    """Test that the report coverage table uses inline CSS"""
    
    # Create a mock table structure for testing
    class MockCell:
        def __init__(self, text):
            self.text = text
    
    class MockRow:
        def __init__(self, cells):
            self.cells = cells
    
    class MockTable:
        def __init__(self, rows):
            self.rows = rows
    
    class MockDocument:
        def __init__(self, tables):
            self.tables = tables
    
    # Create a mock document with a report coverage table
    mock_table = MockTable([
        MockRow([MockCell("Report Attribute"), MockCell("Details")]),
        MockRow([MockCell("Forecast Period"), MockCell("2024-2030")]),
        MockRow([MockCell("Market Size"), MockCell("USD 1.2 Billion")])
    ])
    
    mock_doc = MockDocument([mock_table])
    
    # Test the extraction
    result = extract_report_coverage_table_with_style("mock_path")
    
    print("Testing inline CSS conversion...")
    print("=" * 50)
    
    # Check if the result contains inline styles instead of <style> tags
    if "<style>" in result:
        print("❌ FAIL: Still contains internal CSS (<style> tags)")
        return False
    elif "style=" in result:
        print("✅ PASS: Contains inline CSS (style= attributes)")
        print("\nSample output:")
        print(result[:500] + "..." if len(result) > 500 else result)
        return True
    else:
        print("⚠️  WARNING: No CSS found in output")
        return False

if __name__ == "__main__":
    success = test_inline_css_conversion()
    sys.exit(0 if success else 1)
