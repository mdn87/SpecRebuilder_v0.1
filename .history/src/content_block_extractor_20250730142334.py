#!/usr/bin/env python3
"""
Content Block Extractor

This script extracts content blocks from Word documents, removes blank lines,
and identifies level numbers for each block. It assumes the first, second,
and last blocks are special (section number, title, end of section).
"""

import json
import sys
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from word_to_json import WordToJsonConverter

@dataclass
class ContentBlock:
    """Represents a content block with its level information"""
    text: str
    level_number: Optional[int] = None
    block_type: str = "content"  # "section_number", "section_title", "content", "end_of_section"
    index: int = 0

class ContentBlockExtractor:
    """Extracts content blocks from Word documents"""
    
    def __init__(self):
        self.blocks = []
    
    def extract_content_blocks(self, docx_path: str) -> List[ContentBlock]:
        """Extract content blocks from a Word document"""
        # First convert to JSON
        converter = WordToJsonConverter()
        json_path = converter.convert_to_json(docx_path)
        
        # Load the JSON structure
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Extract non-empty paragraphs
        content_blocks = []
        for paragraph in data.get('paragraphs', []):
            text = paragraph.get('text', '').strip()
            if text:  # Only include non-empty paragraphs
                # Extract level information
                numbering_info = paragraph.get('numbering', {})
                level_number = numbering_info.get('level') if numbering_info else None
                
                block = ContentBlock(
                    text=text,
                    level_number=level_number,
                    index=len(content_blocks)
                )
                content_blocks.append(block)
        
        # Classify blocks based on position
        if len(content_blocks) >= 3:
            # First block: section number
            content_blocks[0].block_type = "section_number"
            
            # Second block: section title
            content_blocks[1].block_type = "section_title"
            
            # Last block: end of section
            content_blocks[-1].block_type = "end_of_section"
            
            # All others are content blocks
            for i in range(2, len(content_blocks) - 1):
                content_blocks[i].block_type = "content"
        
        self.blocks = content_blocks
        return content_blocks
    
    def analyze_level_distribution(self) -> Dict[str, Any]:
        """Analyze the distribution of levels in content blocks"""
        if not self.blocks:
            return {}
        
        # Count blocks by type
        type_counts = {}
        for block in self.blocks:
            block_type = block.block_type
            type_counts[block_type] = type_counts.get(block_type, 0) + 1
        
        # Analyze level numbers for content blocks
        content_blocks = [b for b in self.blocks if b.block_type == "content"]
        level_analysis = {
            'total_content_blocks': len(content_blocks),
            'blocks_with_levels': len([b for b in content_blocks if b.level_number is not None]),
            'blocks_without_levels': len([b for b in content_blocks if b.level_number is None]),
            'level_distribution': {},
            'missing_levels': []
        }
        
        # Count level distribution
        for block in content_blocks:
            if block.level_number is not None:
                level = block.level_number
                level_analysis['level_distribution'][level] = level_analysis['level_distribution'].get(level, 0) + 1
            else:
                level_analysis['missing_levels'].append({
                    'index': block.index,
                    'text': block.text[:50] + "..." if len(block.text) > 50 else block.text
                })
        
        return {
            'block_type_counts': type_counts,
            'level_analysis': level_analysis,
            'total_blocks': len(self.blocks)
        }
    
    def generate_report(self) -> str:
        """Generate a human-readable report"""
        if not self.blocks:
            return "No content blocks found."
        
        analysis = self.analyze_level_distribution()
        
        report = []
        report.append("=== CONTENT BLOCK ANALYSIS ===")
        report.append(f"Total blocks: {analysis['total_blocks']}")
        report.append("")
        
        # Block type summary
        report.append("Block Type Distribution:")
        for block_type, count in analysis['block_type_counts'].items():
            report.append(f"  {block_type}: {count}")
        report.append("")
        
        # Level analysis
        level_analysis = analysis['level_analysis']
        report.append("Content Block Level Analysis:")
        report.append(f"  Total content blocks: {level_analysis['total_content_blocks']}")
        report.append(f"  Blocks with levels: {level_analysis['blocks_with_levels']}")
        report.append(f"  Blocks without levels: {level_analysis['blocks_without_levels']}")
        report.append("")
        
        # Level distribution
        if level_analysis['level_distribution']:
            report.append("Level Distribution:")
            for level, count in sorted(level_analysis['level_distribution'].items()):
                report.append(f"  Level {level}: {count} blocks")
            report.append("")
        
        # Missing levels
        if level_analysis['missing_levels']:
            report.append("Blocks Missing Level Numbers:")
            for missing in level_analysis['missing_levels']:
                report.append(f"  Block {missing['index']}: {missing['text']}")
            report.append("")
        
        # Content block preview
        content_blocks = [b for b in self.blocks if b.block_type == "content"]
        if content_blocks:
            report.append("Content Block Preview (first 5):")
            for i, block in enumerate(content_blocks[:5]):
                level_info = f" (Level {block.level_number})" if block.level_number is not None else " (No level)"
                report.append(f"  {i+1}. {block.text[:60]}{level_info}")
        
        return "\n".join(report)
    
    def save_blocks_to_json(self, output_path: str):
        """Save content blocks to JSON format"""
        if not self.blocks:
            print("No blocks to save.")
            return
        
        # Convert blocks to dictionary format
        blocks_data = []
        for block in self.blocks:
            block_data = {
                'text': block.text,
                'level_number': block.level_number,
                'block_type': block.block_type,
                'index': block.index
            }
            blocks_data.append(block_data)
        
        # Create output data
        output_data = {
            'blocks': blocks_data,
            'analysis': self.analyze_level_distribution(),
            'total_blocks': len(self.blocks)
        }
        
        # Save to JSON
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"Content blocks saved to: {output_path}")
    
    def print_blocks(self, max_blocks: int = 10):
        """Print content blocks for inspection"""
        if not self.blocks:
            print("No content blocks found.")
            return
        
        print(f"=== CONTENT BLOCKS (showing first {max_blocks}) ===")
        for i, block in enumerate(self.blocks[:max_blocks]):
            level_info = f" [Level {block.level_number}]" if block.level_number is not None else " [No level]"
            print(f"{i+1:2d}. [{block.block_type:15}] {block.text[:60]}{level_info}")
        
        if len(self.blocks) > max_blocks:
            print(f"... and {len(self.blocks) - max_blocks} more blocks")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python content_block_extractor.py <docx_file> [output_file]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    extractor = ContentBlockExtractor()
    
    try:
        # Extract content blocks
        print(f"Extracting content blocks from: {docx_path}")
        blocks = extractor.extract_content_blocks(docx_path)
        print(f"Extracted {len(blocks)} content blocks")
        
        # Print analysis
        print("\n" + extractor.generate_report())
        
        # Print first few blocks for inspection
        print("\n")
        extractor.print_blocks(15)
        
        # Save to JSON if output path specified
        if output_path:
            extractor.save_blocks_to_json(output_path)
        else:
            # Save with default name
            base_name = Path(docx_path).stem
            default_output = f"{base_name}_content_blocks.json"
            extractor.save_blocks_to_json(default_output)
        
    except Exception as e:
        print(f"Error extracting content blocks: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 