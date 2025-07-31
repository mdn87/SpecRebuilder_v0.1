#!/usr/bin/env python3
"""
Content Block Pattern Analyzer

This script analyzes content blocks to identify patterns and suggest
level assignments for blocks that are missing level numbers.
"""

import json
import sys
import os
import re
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from content_block_extractor import ContentBlockExtractor, ContentBlock

@dataclass
class PatternMatch:
    """Represents a pattern match for level assignment"""
    pattern: str
    suggested_level: int
    confidence: float
    examples: List[str]

class BlockPatternAnalyzer:
    """Analyzes content blocks to identify patterns and suggest levels"""
    
    def __init__(self):
        self.patterns = []
        self.blocks = []
    
    def load_blocks_from_json(self, json_path: str) -> List[ContentBlock]:
        """Load content blocks from JSON file"""
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        blocks = []
        for block_data in data.get('blocks', []):
            block = ContentBlock(
                text=block_data['text'],
                level_number=block_data['level_number'],
                block_type=block_data['block_type'],
                index=block_data['index']
            )
            blocks.append(block)
        
        self.blocks = blocks
        return blocks
    
    def analyze_text_patterns(self) -> List[PatternMatch]:
        """Analyze text patterns to identify level indicators"""
        patterns = []
        
        # Common specification patterns
        spec_patterns = [
            # Major sections (Level 0)
            (r'^[A-Z\s]+$', 0, 0.9),  # All caps, no numbers
            (r'^\d+\.\d+\s+[A-Z\s]+$', 0, 0.8),  # Numbered sections like "2.01 GENERAL"
            
            # Subsections (Level 1)
            (r'^[A-Z]\.\s+[A-Z]', 1, 0.7),  # A. B. C. style
            (r'^\d+\.\d+\.\d+\s+[A-Z]', 1, 0.7),  # 2.01.01 style
            
            # Items (Level 2)
            (r'^\d+\.\s+[A-Z]', 2, 0.6),  # 1. 2. 3. style
            (r'^[a-z]\.\s+[A-Z]', 2, 0.6),  # a. b. c. style
            
            # Sub-items (Level 3)
            (r'^\(\d+\)\s+[A-Z]', 3, 0.5),  # (1) (2) style
            (r'^[a-z]\)\s+[A-Z]', 3, 0.5),  # a) b) style
            
            # BWA-specific patterns
            (r'^BWA-PART\d*', 0, 0.9),  # BWA-PART
            (r'^BWA-SUBSECTION\d*', 1, 0.8),  # BWA-SUBSECTION
            (r'^BWA-Item\d*', 2, 0.7),  # BWA-Item
            (r'^BWA-List\d*', 3, 0.6),  # BWA-List
            (r'^BWA-SubItem\d*', 4, 0.5),  # BWA-SubItem
            (r'^BWA-SubList\d*', 5, 0.4),  # BWA-SubList
        ]
        
        # Test patterns against all content blocks
        for pattern, suggested_level, confidence in spec_patterns:
            examples = []
            matches = 0
            total_content_blocks = 0
            
            for block in self.blocks:
                if block.block_type == "content":
                    total_content_blocks += 1
                    if re.match(pattern, block.text.strip()):
                        matches += 1
                        examples.append(block.text[:50])
            
            # Calculate confidence based on match rate
            if total_content_blocks > 0:
                actual_confidence = (matches / total_content_blocks) * confidence
                if matches > 0:
                    patterns.append(PatternMatch(
                        pattern=pattern,
                        suggested_level=suggested_level,
                        confidence=actual_confidence,
                        examples=examples[:3]  # Keep first 3 examples
                    ))
        
        # Sort by confidence
        patterns.sort(key=lambda x: x.confidence, reverse=True)
        self.patterns = patterns
        return patterns
    
    def suggest_levels_for_missing_blocks(self) -> List[Dict[str, Any]]:
        """Suggest levels for blocks that don't have them"""
        suggestions = []
        
        for block in self.blocks:
            if block.block_type == "content" and block.level_number is None:
                best_match = None
                best_confidence = 0
                
                for pattern in self.patterns:
                    if re.match(pattern.pattern, block.text.strip()):
                        if pattern.confidence > best_confidence:
                            best_match = pattern
                            best_confidence = pattern.confidence
                
                if best_match:
                    suggestions.append({
                        'block_index': block.index,
                        'text': block.text,
                        'suggested_level': best_match.suggested_level,
                        'confidence': best_confidence,
                        'pattern': best_match.pattern
                    })
                else:
                    suggestions.append({
                        'block_index': block.index,
                        'text': block.text,
                        'suggested_level': None,
                        'confidence': 0,
                        'pattern': None
                    })
        
        return suggestions
    
    def generate_analysis_report(self) -> str:
        """Generate a comprehensive analysis report"""
        if not self.blocks:
            return "No blocks to analyze."
        
        # Analyze patterns
        patterns = self.analyze_text_patterns()
        
        # Get suggestions
        suggestions = self.suggest_levels_for_missing_blocks()
        
        # Count blocks by type
        content_blocks = [b for b in self.blocks if b.block_type == "content"]
        blocks_with_levels = [b for b in content_blocks if b.level_number is not None]
        blocks_without_levels = [b for b in content_blocks if b.level_number is None]
        
        report = []
        report.append("=== CONTENT BLOCK PATTERN ANALYSIS ===")
        report.append(f"Total content blocks: {len(content_blocks)}")
        report.append(f"Blocks with levels: {len(blocks_with_levels)}")
        report.append(f"Blocks without levels: {len(blocks_without_levels)}")
        report.append("")
        
        # Pattern analysis
        if patterns:
            report.append("Identified Patterns:")
            for i, pattern in enumerate(patterns[:10]):  # Show top 10
                report.append(f"  {i+1}. Pattern: {pattern.pattern}")
                report.append(f"     Suggested Level: {pattern.suggested_level}")
                report.append(f"     Confidence: {pattern.confidence:.2f}")
                report.append(f"     Examples: {', '.join(pattern.examples)}")
                report.append("")
        
        # Suggestions
        if suggestions:
            report.append("Level Suggestions for Missing Blocks:")
            for suggestion in suggestions[:10]:  # Show first 10
                if suggestion['suggested_level'] is not None:
                    report.append(f"  Block {suggestion['block_index']}: Level {suggestion['suggested_level']} "
                               f"(confidence: {suggestion['confidence']:.2f})")
                    report.append(f"    Text: {suggestion['text'][:60]}...")
                    report.append(f"    Pattern: {suggestion['pattern']}")
                else:
                    report.append(f"  Block {suggestion['block_index']}: No pattern match")
                    report.append(f"    Text: {suggestion['text'][:60]}...")
                report.append("")
        
        return "\n".join(report)
    
    def save_suggestions_to_json(self, output_path: str):
        """Save level suggestions to JSON"""
        suggestions = self.suggest_levels_for_missing_blocks()
        
        output_data = {
            'suggestions': suggestions,
            'patterns': [
                {
                    'pattern': p.pattern,
                    'suggested_level': p.suggested_level,
                    'confidence': p.confidence,
                    'examples': p.examples
                }
                for p in self.patterns
            ],
            'summary': {
                'total_blocks': len(self.blocks),
                'content_blocks': len([b for b in self.blocks if b.block_type == "content"]),
                'blocks_with_levels': len([b for b in self.blocks if b.block_type == "content" and b.level_number is not None]),
                'blocks_without_levels': len([b for b in self.blocks if b.block_type == "content" and b.level_number is None]),
                'suggestions_made': len([s for s in suggestions if s['suggested_level'] is not None])
            }
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"Suggestions saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python block_pattern_analyzer.py <content_blocks.json> [output_file]")
        sys.exit(1)
    
    json_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(json_path):
        print(f"Error: File not found: {json_path}")
        sys.exit(1)
    
    analyzer = BlockPatternAnalyzer()
    
    try:
        # Load blocks
        print(f"Loading content blocks from: {json_path}")
        blocks = analyzer.load_blocks_from_json(json_path)
        print(f"Loaded {len(blocks)} blocks")
        
        # Generate analysis
        print("\n" + analyzer.generate_analysis_report())
        
        # Save suggestions if output path specified
        if output_path:
            analyzer.save_suggestions_to_json(output_path)
        else:
            # Save with default name
            base_name = Path(json_path).stem
            default_output = f"{base_name}_suggestions.json"
            analyzer.save_suggestions_to_json(default_output)
        
    except Exception as e:
        print(f"Error analyzing patterns: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 