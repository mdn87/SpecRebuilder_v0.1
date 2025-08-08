#!/usr/bin/env python3
"""
Flexible List Analyzer - Context-aware list structure detection
Uses contextual inference and dynamic level assignment instead of hard-coded patterns.
"""

import json
import re
import sys
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, asdict
from enum import Enum

class ListFormat(Enum):
    """Standard OpenXML numbering formats"""
    DECIMAL = "decimal"
    UPPER_LETTER = "upperLetter"
    LOWER_LETTER = "lowerLetter"
    UPPER_ROMAN = "upperRoman"
    LOWER_ROMAN = "lowerRoman"
    BULLET = "bullet"
    NONE = "none"

@dataclass
class ListContext:
    """Context for list level inference"""
    current_level: int
    numbering_stack: List[Tuple[str, int]]  # (numbering_pattern, level)
    indent_stack: List[int]
    format_stack: List[ListFormat]

@dataclass
class FlexibleBlock:
    """Flexible content block with context-aware metadata"""
    index: int
    text: str
    cleaned_content: str
    level: Optional[int]
    num_fmt: Optional[ListFormat]
    list_id: Optional[int]
    numbering_pattern: Optional[str]
    inferred_number: Optional[str]
    is_list_item: bool
    indentation_level: Optional[int]
    context_hints: List[str]
    confidence_score: float
    parent_context: Optional[str]

class FlexibleListAnalyzer:
    """Context-aware list structure analyzer"""
    
    def __init__(self):
        # Numbering format detection patterns (not tied to levels)
        self.format_patterns = {
            ListFormat.DECIMAL: [
                r'^\d+\.',           # 1., 2., 3.
                r'^\d+\.\d+',        # 1.0, 1.1, 2.0
                r'^\d+\.\d+\.\d+',  # 1.0.1, 1.1.2
            ],
            ListFormat.UPPER_LETTER: [
                r'^[A-Z]\.',         # A., B., C.
                r'^[A-Z]\.\d+',      # A.1, B.2
            ],
            ListFormat.LOWER_LETTER: [
                r'^[a-z]\.',         # a., b., c.
                r'^[a-z]\.\d+',      # a.1, b.2
            ],
            ListFormat.UPPER_ROMAN: [
                r'^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.',  # I., II., III.
            ],
            ListFormat.LOWER_ROMAN: [
                r'^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\.',  # i., ii., iii.
            ]
        }
    
    def analyze_document(self, json_path: str) -> Dict[str, Any]:
        """Main analysis function with flexible level assignment"""
        print(f"Loading analysis from: {json_path}")
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Extract paragraphs
        paragraphs = data.get('sample_paragraphs', [])
        
        # Convert to flexible blocks
        flexible_blocks = self._create_flexible_blocks(paragraphs)
        
        # Perform context-aware level assignment
        self._assign_levels_contextually(flexible_blocks)
        
        # Group lists based on context
        list_groups = self._group_lists_contextually(flexible_blocks)
        
        # Calculate confidence scores
        for block in flexible_blocks:
            block.confidence_score = self._calculate_confidence(block)
        
        # Generate comprehensive report
        report = self._generate_report(flexible_blocks, list_groups, data)
        
        return report
    
    def _create_flexible_blocks(self, paragraphs: List[Dict]) -> List[FlexibleBlock]:
        """Convert raw paragraphs to flexible blocks"""
        flexible_blocks = []
        
        for para in paragraphs:
            text = para.get('text', '')
            numbering_pattern = para.get('list_number', '')
            inferred_number = para.get('inferred_number')
            
            # Determine if this is a list item
            is_list_item = bool(numbering_pattern or inferred_number)
            
            # Clean content
            cleaned_content = self._clean_content(text, numbering_pattern, inferred_number)
            
            # Detect numbering format (not tied to level)
            num_fmt = self._detect_numbering_format(numbering_pattern or inferred_number or '')
            
            # Create flexible block
            block = FlexibleBlock(
                index=para.get('index', 0),
                text=text,
                cleaned_content=cleaned_content,
                level=None,  # Will be assigned contextually
                num_fmt=num_fmt,
                list_id=None,  # Will be assigned later
                numbering_pattern=numbering_pattern or inferred_number,
                inferred_number=inferred_number,
                is_list_item=is_list_item,
                indentation_level=None,
                context_hints=[],
                confidence_score=0.0,
                parent_context=None
            )
            
            flexible_blocks.append(block)
        
        return flexible_blocks
    
    def _clean_content(self, text: str, numbering_pattern: str, inferred_number: Optional[str]) -> str:
        """Remove numbering prefixes and clean content"""
        if numbering_pattern:
            pattern = re.escape(numbering_pattern)
            cleaned = re.sub(f'^{pattern}\\s*\\t?\\s*', '', text)
            return cleaned.strip()
        elif inferred_number:
            pattern = re.escape(inferred_number)
            cleaned = re.sub(f'^{pattern}\\s*\\t?\\s*', '', text)
            return cleaned.strip()
        else:
            return text.strip()
    
    def _detect_numbering_format(self, numbering: str) -> Optional[ListFormat]:
        """Detect numbering format without level assumptions"""
        if not numbering:
            return None
        
        for fmt, patterns in self.format_patterns.items():
            for pattern in patterns:
                if re.match(pattern, numbering):
                    return fmt
        
        return None
    
    def _assign_levels_contextually(self, blocks: List[FlexibleBlock]):
        """Assign levels based on context and structure, not hard-coded patterns"""
        context = ListContext(
            current_level=0,
            numbering_stack=[],
            indent_stack=[],
            format_stack=[]
        )
        
        for i, block in enumerate(blocks):
            if not block.is_list_item:
                # Non-list item resets context
                context.numbering_stack = []
                context.indent_stack = []
                context.format_stack = []
                context.current_level = 0
                continue
            
            # Analyze this block's context
            level = self._infer_level_from_context(block, context, blocks[:i])
            block.level = level
            
            # Update context
            if block.numbering_pattern:
                context.numbering_stack.append((block.numbering_pattern, level))
                if block.num_fmt:
                    context.format_stack.append(block.num_fmt)
                context.current_level = level
    
    def _infer_level_from_context(self, block: FlexibleBlock, context: ListContext, previous_blocks: List[FlexibleBlock]) -> int:
        """Infer level based on context and previous items"""
        numbering = block.numbering_pattern or ''
        
        # If this is the first list item, start at level 0
        if not context.numbering_stack:
            return 0
        
        # Analyze relationship to previous items
        for prev_numbering, prev_level in reversed(context.numbering_stack):
            relationship = self._analyze_numbering_relationship(numbering, prev_numbering)
            
            if relationship == 'sublevel':
                return prev_level + 1
            elif relationship == 'sibling':
                return prev_level
            elif relationship == 'parent':
                return max(0, prev_level - 1)
            elif relationship == 'new_list':
                return 0
        
        # Default: assume same level as previous
        return context.current_level
    
    def _analyze_numbering_relationship(self, current: str, previous: str) -> str:
        """Analyze the relationship between two numbering patterns"""
        if not current or not previous:
            return 'new_list'
        
        # Extract base patterns
        current_base = self._extract_base_pattern(current)
        previous_base = self._extract_base_pattern(previous)
        
        # Check for sublevel patterns
        if self._is_sublevel_pattern(current, previous):
            return 'sublevel'
        
        # Check for sibling patterns (same format, different number)
        if self._is_sibling_pattern(current, previous):
            return 'sibling'
        
        # Check for parent patterns (less indented, different format)
        if self._is_parent_pattern(current, previous):
            return 'parent'
        
        # Default: new list
        return 'new_list'
    
    def _extract_base_pattern(self, numbering: str) -> str:
        """Extract the base pattern from numbering"""
        # Remove numbers but keep format
        if re.match(r'^\d+\.', numbering):
            return 'decimal'
        elif re.match(r'^[A-Z]\.', numbering):
            return 'upper_letter'
        elif re.match(r'^[a-z]\.', numbering):
            return 'lower_letter'
        elif re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.', numbering):
            return 'upper_roman'
        elif re.match(r'^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\.', numbering):
            return 'lower_roman'
        else:
            return 'unknown'
    
    def _is_sublevel_pattern(self, current: str, previous: str) -> bool:
        """Check if current is a sublevel of previous"""
        # Common sublevel patterns
        sublevel_patterns = [
            # Decimal sublevels
            (r'^\d+\.', r'^\d+\.\d+'),  # 1. -> 1.1
            (r'^\d+\.\d+', r'^\d+\.\d+\.\d+'),  # 1.1 -> 1.1.1
            
            # Letter sublevels
            (r'^\d+\.', r'^[A-Z]\.'),  # 1. -> A.
            (r'^[A-Z]\.', r'^\d+\.'),  # A. -> 1.
            (r'^\d+\.', r'^[a-z]\.'),  # 1. -> a.
            (r'^[A-Z]\.', r'^[a-z]\.'),  # A. -> a.
            
            # Roman sublevels
            (r'^[A-Z]\.', r'^(i|ii|iii)\.'),  # A. -> i.
            (r'^\d+\.', r'^(i|ii|iii)\.'),  # 1. -> i.
        ]
        
        for prev_pattern, curr_pattern in sublevel_patterns:
            if re.match(prev_pattern, previous) and re.match(curr_pattern, current):
                return True
        
        return False
    
    def _is_sibling_pattern(self, current: str, previous: str) -> bool:
        """Check if current is a sibling of previous"""
        current_base = self._extract_base_pattern(current)
        previous_base = self._extract_base_pattern(previous)
        
        return current_base == previous_base
    
    def _is_parent_pattern(self, current: str, previous: str) -> bool:
        """Check if current is a parent of previous"""
        # This is more complex and might need additional context
        # For now, assume different formats at same level are new lists
        current_base = self._extract_base_pattern(current)
        previous_base = self._extract_base_pattern(previous)
        
        return current_base != previous_base
    
    def _group_lists_contextually(self, blocks: List[FlexibleBlock]) -> List[List[int]]:
        """Group lists based on context and level continuity"""
        list_groups = []
        current_group = []
        current_level = None
        
        for i, block in enumerate(blocks):
            if block.is_list_item and block.level is not None:
                # Start new group if level changes significantly
                if current_level is None or abs(block.level - current_level) <= 1:
                    if not current_group:
                        current_group = [i]
                    else:
                        current_group.append(i)
                    current_level = block.level
                else:
                    # Level change too big, start new group
                    if current_group:
                        list_groups.append(current_group)
                    current_group = [i]
                    current_level = block.level
            else:
                # Non-list item ends current group
                if current_group:
                    list_groups.append(current_group)
                    current_group = []
                    current_level = None
        
        # Don't forget the last group
        if current_group:
            list_groups.append(current_group)
        
        # Assign list IDs
        for group_id, group in enumerate(list_groups, 1):
            for block_idx in group:
                blocks[block_idx].list_id = group_id
        
        return list_groups
    
    def _calculate_confidence(self, block: FlexibleBlock) -> float:
        """Calculate confidence score for the analysis"""
        confidence = 0.0
        
        if block.level is not None:
            confidence += 0.3
        
        if block.num_fmt is not None:
            confidence += 0.3
        
        if block.list_id is not None:
            confidence += 0.2
        
        if block.numbering_pattern:
            confidence += 0.2
        
        return min(confidence, 1.0)
    
    def _generate_report(self, blocks: List[FlexibleBlock], list_groups: List[List[int]], original_data: Dict) -> Dict[str, Any]:
        """Generate comprehensive analysis report"""
        
        # Convert blocks to serializable format
        serializable_blocks = []
        for block in blocks:
            block_dict = asdict(block)
            # Convert enum to string
            if block_dict['num_fmt']:
                block_dict['num_fmt'] = block_dict['num_fmt'].value
            serializable_blocks.append(block_dict)
        
        # Statistics
        total_blocks = len(blocks)
        list_items = [b for b in blocks if b.is_list_item]
        non_list_items = [b for b in blocks if not b.is_list_item]
        
        level_distribution = {}
        format_distribution = {}
        
        for block in list_items:
            if block.level is not None:
                level_distribution[block.level] = level_distribution.get(block.level, 0) + 1
            
            if block.num_fmt is not None:
                fmt = block.num_fmt.value
                format_distribution[fmt] = format_distribution.get(fmt, 0) + 1
        
        report = {
            "flexible_analysis": {
                "total_blocks": total_blocks,
                "list_items": len(list_items),
                "non_list_items": len(non_list_items),
                "list_groups": len(list_groups),
                "level_distribution": level_distribution,
                "format_distribution": format_distribution,
                "average_confidence": sum(b.confidence_score for b in blocks) / len(blocks) if blocks else 0
            },
            "flexible_blocks": serializable_blocks,
            "list_groups": list_groups,
            "original_analysis": original_data
        }
        
        return report

def main():
    """Main function"""
    if len(sys.argv) != 2:
        print("Usage: python flexible_list_analyzer.py <json_file>")
        sys.exit(1)
    
    json_path = sys.argv[1]
    
    analyzer = FlexibleListAnalyzer()
    report = analyzer.analyze_document(json_path)
    
    # Save flexible analysis
    output_path = json_path.replace('.json', '_flexible.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    
    print(f"Flexible analysis saved to: {output_path}")
    
    # Print summary
    flexible = report['flexible_analysis']
    print(f"\nFlexible Analysis Summary:")
    print(f"Total blocks: {flexible['total_blocks']}")
    print(f"List items: {flexible['list_items']}")
    print(f"List groups: {flexible['list_groups']}")
    print(f"Level distribution: {flexible['level_distribution']}")
    print(f"Format distribution: {flexible['format_distribution']}")
    print(f"Average confidence: {flexible['average_confidence']:.2f}")

if __name__ == "__main__":
    main()
