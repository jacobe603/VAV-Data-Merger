#!/usr/bin/env python3
import os
from app_fixed import read_tw2_data_safe

TW2_FILE = r'C:\Users\Jacob\Claude\VAV2\936290 - UND Flight Operations.tw2'

result = read_tw2_data_safe(TW2_FILE)

if result['success']:
    print(f"Records: {result['row_count']}")
    print(f"Columns: {len(result['columns'])}")
    
    print("\nSample Tag values:")
    tags = []
    for record in result['data']:
        tag = record.get('Tag')
        if tag:
            tags.append(str(tag))
    
    for i, tag in enumerate(sorted(set(tags))[:15]):
        print(f"{i+1:2d}. {tag}")
        
else:
    print(f"Error: {result['error']}")