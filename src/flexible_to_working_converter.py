#!/usr/bin/env python3
"""
Flexible to Working Converter - Converts flexible analysis to working C# rebuilder format
"""

import json
import sys
from typing import Dict, List

def convert_flexible_to_working(flexible_json_path: str, output_path: str):
    """Convert flexible analysis to working C# rebuilder format"""
    print(f"Loading flexible analysis from: {flexible_json_path}")
    
    with open(flexible_json_path, 'r', encoding='utf-8') as f:
        flexible_data = json.load(f)
    
    flexible_blocks = flexible_data.get('flexible_blocks', [])
    
    # Convert to working format (same as the existing working rebuilder expects)
    working_paragraphs = []
    
    for block in flexible_blocks:
        # Create working paragraph format
        working_para = {
            'index': block['index'],
            'list_number': block.get('numbering_pattern', '') or '',
            'inferred_number': block.get('inferred_number'),
            'text': block.get('cleaned_content', block.get('text', '')),  # Use cleaned content
            'combined': block.get('text', ''),  # Keep original combined text
            'level': block.get('level'),
            'deduction_method': None
        }
        
        working_paragraphs.append(working_para)
    
    # Create working format
    working_data = {
        'all_paragraphs': working_paragraphs,  # Match working C# rebuilder expectation
        'flexible_analysis': flexible_data.get('flexible_analysis', {}),
        'list_groups': flexible_data.get('list_groups', [])
    }
    
    # Save working format
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(working_data, f, indent=2, ensure_ascii=False)
    
    print(f"Converted to working format: {output_path}")
    
    # Print summary
    list_items = sum(1 for p in working_paragraphs if p.get('list_number', ''))
    print(f"Total paragraphs: {len(working_paragraphs)}")
    print(f"List items: {list_items}")
    print(f"Levels found: {set(p['level'] for p in working_paragraphs if p['level'] is not None)}")

def main():
    if len(sys.argv) != 3:
        print("Usage: python flexible_to_working_converter.py <flexible_json> <output_json>")
        sys.exit(1)
    
    flexible_path = sys.argv[1]
    output_path = sys.argv[2]
    
    convert_flexible_to_working(flexible_path, output_path)

if __name__ == "__main__":
    main()
