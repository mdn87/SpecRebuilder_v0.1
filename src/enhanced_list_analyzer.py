#!/usr/bin/env python3
"""
Enhanced List Analyzer - Comprehensive list structure detection and normalization
Addresses the core issues with level assignment, list type detection, and grouping.
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
class EnhancedBlock:
    """Enhanced content block with comprehensive list metadata"""
    index: int
    text: str
    cleaned_content: str
    level: Optional[int]
    num_fmt: Optional[ListFormat]
    list_id: Optional[int]
    parent_list_id: Optional[int]
    numbering_pattern: Optional[str]
    inferred_number: Optional[str]
    is_list_item: bool
    is_continuation: bool
    continuation_of: Optional[int]
    indentation_level: Optional[int]
    context_hints: List[str]
    confidence_score: float

class EnhancedListAnalyzer:
    """Comprehensive list structure analyzer and normalizer"""
    
    def __init__(self):
        # Numbering pattern detection regexes
        self.patterns = {
            ListFormat.DECIMAL: [
                r'^(\d+)\.',           # 1., 2., 3.
                r'^(\d+)\.(\d+)',      # 1.0, 1.1, 2.0
                r'^(\d+)\.(\d+)\.(\d+)', # 1.0.1, 1.1.2
            ],
            ListFormat.UPPER_LETTER: [
                r'^([A-Z])\.',         # A., B., C.
                r'^([A-Z])\.(\d+)',    # A.1, B.2
            ],
            ListFormat.LOWER_LETTER: [
                r'^([a-z])\.',         # a., b., c.
                r'^([a-z])\.(\d+)',    # a.1, b.2
            ],
            ListFormat.UPPER_ROMAN: [
                r'^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.',  # I., II., III.
            ],
            ListFormat.LOWER_ROMAN: [
                r'^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\.',  # i., ii., iii.
            ]
        }
        
        # Level inference patterns based on numbering
        self.level_patterns = {
            # Level 0: Main sections (1.0, 2.0, etc.)
            r'^\d+\.0\s': 0,
            # Level 1: Subsections (1.01, 1.02, etc.)
            r'^\d+\.\d{2}\s': 1,
            # Level 2: Items (A., B., C.)
            r'^[A-Z]\.\s': 2,
            # Level 3: Sub-items (1., 2., 3.)
            r'^\d+\.\s': 3,
            # Level 4: Sub-sub-items (a., b., c.)
            r'^[a-z]\.\s': 4,
            # Level 5: Roman numerals (i., ii., iii.)
            r'^[ivx]+\.\s': 5,
        }
    
    def analyze_document(self, json_path: str) -> Dict[str, Any]:
        """Main analysis function"""
        print(f"Loading analysis from: {json_path}")
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Extract paragraphs
        paragraphs = data.get('sample_paragraphs', [])
        
        # Enhanced analysis
        enhanced_blocks = self._enhance_blocks(paragraphs)
        list_groups = self._group_lists(enhanced_blocks)
        level_assignments = self._assign_levels(enhanced_blocks, list_groups)
        final_blocks = self._finalize_analysis(enhanced_blocks, level_assignments)
        
        # Generate comprehensive report
        report = self._generate_report(final_blocks, list_groups, data)
        
        return report
    
    def _enhance_blocks(self, paragraphs: List[Dict]) -> List[EnhancedBlock]:
        """Convert raw paragraphs to enhanced blocks with initial analysis"""
        enhanced_blocks = []
        
        for para in paragraphs:
            text = para.get('text', '')
            combined = para.get('combined', text)
            
            # Extract numbering pattern
            numbering_pattern = para.get('list_number', '')
            inferred_number = para.get('inferred_number')
            
            # Determine if this is a list item
            is_list_item = bool(numbering_pattern or inferred_number)
            
            # Clean content (remove numbering prefix)
            cleaned_content = self._clean_content(text, numbering_pattern, inferred_number)
            
            # Initial level assignment
            level = para.get('level')
            
            # Detect numbering format
            num_fmt = self._detect_numbering_format(numbering_pattern or inferred_number or '')
            
            # Create enhanced block
            block = EnhancedBlock(
                index=para.get('index', 0),
                text=text,
                cleaned_content=cleaned_content,
                level=level,
                num_fmt=num_fmt,
                list_id=None,  # Will be assigned later
                parent_list_id=None,
                numbering_pattern=numbering_pattern or inferred_number,
                inferred_number=inferred_number,
                is_list_item=is_list_item,
                is_continuation=False,
                continuation_of=None,
                indentation_level=None,
                context_hints=[],
                confidence_score=0.0
            )
            
            enhanced_blocks.append(block)
        
        return enhanced_blocks
    
    def _clean_content(self, text: str, numbering_pattern: str, inferred_number: Optional[str]) -> str:
        """Remove numbering prefixes and clean content"""
        if numbering_pattern:
            # Remove the numbering pattern and any following tabs/spaces
            pattern = re.escape(numbering_pattern)
            cleaned = re.sub(f'^{pattern}\\s*\\t?\\s*', '', text)
            return cleaned.strip()
        elif inferred_number:
            # Remove inferred numbering
            pattern = re.escape(inferred_number)
            cleaned = re.sub(f'^{pattern}\\s*\\t?\\s*', '', text)
            return cleaned.strip()
        else:
            return text.strip()
    
    def _detect_numbering_format(self, numbering: str) -> Optional[ListFormat]:
        """Detect the numbering format from the pattern"""
        if not numbering:
            return None
        
        for fmt, patterns in self.patterns.items():
            for pattern in patterns:
                if re.match(pattern, numbering):
                    return fmt
        
        return None
    
    def _group_lists(self, blocks: List[EnhancedBlock]) -> List[List[int]]:
        """Group contiguous list items into lists"""
        list_groups = []
        current_group = []
        
        for i, block in enumerate(blocks):
            if block.is_list_item:
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
        
        return list_groups
    
    def _assign_levels(self, blocks: List[EnhancedBlock], list_groups: List[List[int]]) -> Dict[int, int]:
        """Assign levels to blocks based on numbering patterns and context"""
        level_assignments = {}
        
        for group in list_groups:
            # Analyze the group to determine level hierarchy
            group_levels = self._analyze_group_levels([blocks[i] for i in group])
            
            # Assign levels to blocks in this group
            for i, level in zip(group, group_levels):
                level_assignments[i] = level
        
        return level_assignments
    
    def _analyze_group_levels(self, group_blocks: List[EnhancedBlock]) -> List[int]:
        """Analyze a group of blocks to determine their levels"""
        levels = []
        
        for block in group_blocks:
            numbering = block.numbering_pattern or ''
            
            # Try to infer level from numbering pattern
            level = self._infer_level_from_numbering(numbering)
            
            if level is not None:
                levels.append(level)
            else:
                # Fallback: use existing level or default
                levels.append(block.level or 0)
        
        return levels
    
    def _infer_level_from_numbering(self, numbering: str) -> Optional[int]:
        """Infer level from numbering pattern"""
        for pattern, level in self.level_patterns.items():
            if re.match(pattern, numbering):
                return level
        
        return None
    
    def _finalize_analysis(self, blocks: List[EnhancedBlock], level_assignments: Dict[int, int]) -> List[EnhancedBlock]:
        """Finalize the analysis with all assignments"""
        # Assign list IDs
        list_id = 0
        current_list_blocks = []
        
        for i, block in enumerate(blocks):
            if block.is_list_item:
                if not current_list_blocks:
                    list_id += 1
                current_list_blocks.append(i)
                block.list_id = list_id
                
                # Assign level
                if i in level_assignments:
                    block.level = level_assignments[i]
            else:
                # Non-list item ends current list
                current_list_blocks = []
        
        # Assign parent list IDs for nested lists
        self._assign_parent_list_ids(blocks)
        
        # Calculate confidence scores
        for block in blocks:
            block.confidence_score = self._calculate_confidence(block)
        
        return blocks
    
    def _assign_parent_list_ids(self, blocks: List[EnhancedBlock]):
        """Assign parent list IDs for nested lists"""
        # This is a simplified version - in practice, you'd need more sophisticated logic
        # to detect parent-child relationships based on level hierarchy
        pass
    
    def _calculate_confidence(self, block: EnhancedBlock) -> float:
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
    
    def _generate_report(self, blocks: List[EnhancedBlock], list_groups: List[List[int]], original_data: Dict) -> Dict[str, Any]:
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
            "enhanced_analysis": {
                "total_blocks": total_blocks,
                "list_items": len(list_items),
                "non_list_items": len(non_list_items),
                "list_groups": len(list_groups),
                "level_distribution": level_distribution,
                "format_distribution": format_distribution,
                "average_confidence": sum(b.confidence_score for b in blocks) / len(blocks) if blocks else 0
            },
            "enhanced_blocks": serializable_blocks,
            "list_groups": list_groups,
            "original_analysis": original_data
        }
        
        return report

def main():
    """Main function"""
    if len(sys.argv) != 2:
        print("Usage: python enhanced_list_analyzer.py <json_file>")
        sys.exit(1)
    
    json_path = sys.argv[1]
    
    analyzer = EnhancedListAnalyzer()
    report = analyzer.analyze_document(json_path)
    
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
    print(f"Average confidence: {enhanced['average_confidence']:.2f}")

if __name__ == "__main__":
    main()
