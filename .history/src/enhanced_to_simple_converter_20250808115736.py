#!/usr/bin/env python3
"""
Enhanced to Simple Converter - Converts enhanced analysis to simple format for C# rebuilder
"""

import json
import sys
from typing import Dict, List

def convert_enhanced_to_simple(enhanced_json_path: str, output_path: str):
    """Convert enhanced analysis to simple format"""
    print(f"Loading enhanced analysis from: {enhanced_json_path}")
    
    with open(enhanced_json_path, 'r', encoding='utf-8') as f:
        enhanced_data = json.load(f)
    
    enhanced_blocks = enhanced_data.get('enhanced_blocks', [])
    
    # Convert to simple format
    simple_paragraphs = []
    
    for block in enhanced_blocks:
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
        'sample_paragraphs': simple_paragraphs,  # Changed from 'paragraphs' to 'sample_paragraphs'
        'enhanced_analysis': enhanced_data.get('enhanced_analysis', {}),
        'list_groups': enhanced_data.get('list_groups', [])
    }
    
    # Save simple format
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(simple_data, f, indent=2, ensure_ascii=False)
    
    print(f"Converted to simple format: {output_path}")
    
    # Print summary
    list_items = sum(1 for p in simple_paragraphs if p['is_list_item'])
    print(f"Total paragraphs: {len(simple_paragraphs)}")
    print(f"List items: {list_items}")
    print(f"Levels found: {set(p['level'] for p in simple_paragraphs if p['level'] is not None)}")

def main():
    if len(sys.argv) != 3:
        print("Usage: python enhanced_to_simple_converter.py <enhanced_json> <output_json>")
        sys.exit(1)
    
    enhanced_path = sys.argv[1]
    output_path = sys.argv[2]
    
    convert_enhanced_to_simple(enhanced_path, output_path)

if __name__ == "__main__":
    main()
