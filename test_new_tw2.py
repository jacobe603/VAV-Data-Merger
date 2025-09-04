#!/usr/bin/env python3
"""
Test script to verify the updated TW2 file works with our implementation
"""

import os
from app_fixed import read_tw2_data_safe

# Test the updated TW2 file
TW2_FILE = r'C:\Users\Jacob\Claude\VAV2\936290 - UND Flight Operations.tw2'

print("Testing updated TW2 file...")
print(f"File path: {TW2_FILE}")
print(f"File exists: {os.path.exists(TW2_FILE)}")

if os.path.exists(TW2_FILE):
    result = read_tw2_data_safe(TW2_FILE)
    
    if result['success']:
        print(f"✓ Successfully read TW2 file!")
        print(f"✓ Records found: {result['row_count']}")
        print(f"✓ Columns found: {len(result['columns'])}")
        
        # Show sample data focusing on Tag field
        print(f"\n--- Sample Tag values (first 10) ---")
        for i, record in enumerate(result['data'][:10]):
            tag = record.get('Tag', 'N/A')
            print(f"{i+1:2d}. Tag: {tag}")
            
        print(f"\n--- All unique Tag values ---")
        tags = set()
        for record in result['data']:
            tag = record.get('Tag')
            if tag:
                tags.add(str(tag))
        
        sorted_tags = sorted(tags)
        for tag in sorted_tags:
            print(f"  - {tag}")
            
    else:
        print(f"✗ Error reading TW2 file: {result['error']}")
else:
    print("✗ TW2 file not found!")