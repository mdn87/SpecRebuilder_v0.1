#!/usr/bin/env python3
"""
Simple Enhanced List Analyzer - Focus on core level assignment and list grouping
"""

import json
import re
import sys
from typing import Dict, List, Optional

def analyze_enhanced_structure(json_path: str):
    """Analyze and enhance the list structure"""
    print(f"Loading analysis from: {json_path}")
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    paragraphs = data.get('sample_paragraphs', [])
    enhanced_blocks = []
    
    # Group contiguous list items
    list_groups = []
    current_group = []
    
    for i, para in enumerate(paragraphs):
        # Determine if this is a list item
        numbering = para.get('list_number', '') or para.get('inferred_number', '')
        is_list_item = bool(numbering)
        
        # Clean content
        cleaned_content = clean_content(para.get('text', ''), numbering)
        
        # Enhanced block
        enhanced_block = {
            'index': i,
            'text': para.get('text', ''),
            'cleaned_content': cleaned_content,
            'level': para.get('level'),
            'num_fmt': detect_numbering_format(numbering),
            'list_id': None,  # Will be assigned
            'numbering_pattern': numbering,
            'is_list_item': is_list_item,
            'is_continuation': False,
            'confidence_score': 0.0
        }
        
        enhanced_blocks.append(enhanced_block)
        
        # Group list items
        if is_list_item:
            if not current_group:
                current_group = [i]
            else:
                current_group.append(i)
        else:
            if current_group:
                list_groups.append(current_group)
                current_group = []
    
    # Don't forget the last group
    if current_group:
        list_groups.append(current_group)
    
    # Assign list IDs
    for group_id, group in enumerate(list_groups, 1):
        for block_idx in group:
            enhanced_blocks[block_idx]['list_id'] = group_id
    
    # Assign levels based on numbering patterns
    assign_levels(enhanced_blocks, list_groups)
    
    # Calculate confidence scores
    for block in enhanced_blocks:
        block['confidence_score'] = calculate_confidence(block)
    
    # Generate report
    report = {
        'enhanced_analysis': {
            'total_blocks': len(enhanced_blocks),
            'list_items': sum(1 for b in enhanced_blocks if b['is_list_item']),
            'list_groups': len(list_groups),
            'level_distribution': get_level_distribution(enhanced_blocks),
            'format_distribution': get_format_distribution(enhanced_blocks)
        },
        'enhanced_blocks': enhanced_blocks,
        'list_groups': list_groups,
        'original_analysis': data
    }
    
    return report

def clean_content(text: str, numbering: str) -> str:
    """Remove numbering prefixes and clean content"""
    if numbering:
        # Remove the numbering pattern and any following tabs/spaces
        pattern = re.escape(numbering)
        cleaned = re.sub(f'^{pattern}\\s*\\t?\\s*', '', text)
        return cleaned.strip()
    else:
        return text.strip()

def detect_numbering_format(numbering: str) -> Optional[str]:
    """Detect the numbering format from the pattern"""
    if not numbering:
        return None
    
    patterns = {
        'decimal': [r'^\d+\.', r'^\d+\.\d+', r'^\d+\.\d+\.\d+'],
        'upperLetter': [r'^[A-Z]\.', r'^[A-Z]\.\d+'],
        'lowerLetter': [r'^[a-z]\.', r'^[a-z]\.\d+'],
        'upperRoman': [r'^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.'],
        'lowerRoman': [r'^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\.']
    }
    
    for fmt, pattern_list in patterns.items():
        for pattern in pattern_list:
            if re.match(pattern, numbering):
                return fmt
    
    return None

def assign_levels(blocks: List[Dict], list_groups: List[List[int]]):
    """Assign levels to blocks based on numbering patterns"""
    level_patterns = {
        r'^\d+\.0\s': 0,      # 1.0, 2.0
        r'^\d+\.\d{2}\s': 1,  # 1.01, 1.02
        r'^[A-Z]\.\s': 2,     # A., B., C.
        r'^\d+\.\s': 3,       # 1., 2., 3.
        r'^[a-z]\.\s': 4,     # a., b., c.
        r'^[ivx]+\.\s': 5,    # i., ii., iii.
    }
    
    for group in list_groups:
        for block_idx in group:
            block = blocks[block_idx]
            numbering = block['numbering_pattern']
            
            # Try to infer level from numbering pattern
            level = None
            for pattern, level_val in level_patterns.items():
                if re.match(pattern, numbering):
                    level = level_val
                    break
            
            if level is not None:
                block['level'] = level
            elif block['level'] is None:
                block['level'] = 0  # Default level

def calculate_confidence(block: Dict) -> float:
    """Calculate confidence score for the analysis"""
    confidence = 0.0
    
    if block['level'] is not None:
        confidence += 0.3
    
    if block['num_fmt'] is not None:
        confidence += 0.3
    
    if block['list_id'] is not None:
        confidence += 0.2
    
    if block['numbering_pattern']:
        confidence += 0.2
    
    return min(confidence, 1.0)

def get_level_distribution(blocks: List[Dict]) -> Dict[int, int]:
    """Get distribution of levels"""
    distribution = {}
    for block in blocks:
        if block['is_list_item'] and block['level'] is not None:
            level = block['level']
            distribution[level] = distribution.get(level, 0) + 1
    return distribution

def get_format_distribution(blocks: List[Dict]) -> Dict[str, int]:
    """Get distribution of numbering formats"""
    distribution = {}
    for block in blocks:
        if block['is_list_item'] and block['num_fmt'] is not None:
            fmt = block['num_fmt']
            distribution[fmt] = distribution.get(fmt, 0) + 1
    return distribution

def main():
    if len(sys.argv) != 2:
        print("Usage: python simple_enhanced_analyzer.py <json_file>")
        sys.exit(1)
    
    json_path = sys.argv[1]
    report = analyze_enhanced_structure(json_path)
    
    # Save enhanced analysis
    output_path = json_path.replace('.json', '_enhanced.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    
    print(f"Enhanced analysis saved to: {output_path}")
    
    # Print summary
    enhanced = report['enhanced_analysis']
    print(f"\nEnhanced Analysis Summary:")
    print(f"Total blocks: {enhanced['total_blocks']}")
    print(f"List items: {enhanced['list_items']}")
    print(f"List groups: {enhanced['list_groups']}")
    print(f"Level distribution: {enhanced['level_distribution']}")
    print(f"Format distribution: {enhanced['format_distribution']}")

if __name__ == "__main__":
    main()
