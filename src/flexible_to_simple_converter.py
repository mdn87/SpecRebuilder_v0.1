#!/usr/bin/env python3
"""
Flexible to Simple Converter - Converts flexible analysis to simple format for C# rebuilder
"""

import json
import sys
from typing import Dict, List

def convert_flexible_to_simple(flexible_json_path: str, output_path: str):
    """Convert flexible analysis to simple format"""
    print(f"Loading flexible analysis from: {flexible_json_path}")
    
    with open(flexible_json_path, 'r', encoding='utf-8') as f:
        flexible_data = json.load(f)
    
    flexible_blocks = flexible_data.get('flexible_blocks', [])
    
    # Convert to simple format
    simple_paragraphs = []
    
    for block in flexible_blocks:
        # Create simple paragraph format matching the original structure
        simple_para = {
            'index': block['index'],
            'list_number': block.get('numbering_pattern', '') or '',
            'inferred_number': block.get('inferred_number'),
            'text': block.get('cleaned_content', block.get('text', '')),  # Use cleaned content
            'combined': block.get('text', ''),  # Keep original combined text
            'level': block.get('level'),
            'deduction_method': None
        }
        
        simple_paragraphs.append(simple_para)
    
    # Create simple format
    simple_data = {
        'all_paragraphs': simple_paragraphs,  # Match C# rebuilder expectation
        'flexible_analysis': flexible_data.get('flexible_analysis', {}),
        'list_groups': flexible_data.get('list_groups', [])
    }
    
    # Save simple format
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(simple_data, f, indent=2, ensure_ascii=False)
    
    print(f"Converted to simple format: {output_path}")
    
    # Print summary
    list_items = sum(1 for p in simple_paragraphs if p.get('list_number', ''))
    print(f"Total paragraphs: {len(simple_paragraphs)}")
    print(f"List items: {list_items}")
    print(f"Levels found: {set(p['level'] for p in simple_paragraphs if p['level'] is not None)}")

def main():
    if len(sys.argv) != 3:
        print("Usage: python flexible_to_simple_converter.py <flexible_json> <output_json>")
        sys.exit(1)
    
    flexible_path = sys.argv[1]
    output_path = sys.argv[2]
    
    convert_flexible_to_simple(flexible_path, output_path)

if __name__ == "__main__":
    main()
