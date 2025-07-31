#!/usr/bin/env python3
"""
Simple Content Block Extractor

This script extracts content blocks from Word documents, removes blank lines,
and captures any existing list level data without analysis or processing.
"""

import json
import sys
import os
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from word_to_json import WordToJsonConverter

@dataclass
class ContentBlock:
    """Represents a content block with its level information"""
    text: str
    level_number: Optional[int] = None
    numbering_id: Optional[int] = None
    block_type: str = "content"  # "section_number", "section_title", "content", "end_of_section"
    index: int = 0

class SimpleContentExtractor:
    """Extracts content blocks from Word documents without analysis"""
    
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
                # Extract existing numbering information (no processing)
                numbering_info = paragraph.get('numbering', {})
                level_number = numbering_info.get('level') if numbering_info else None
                numbering_id = numbering_info.get('id') if numbering_info else None
                
                block = ContentBlock(
                    text=text,
                    level_number=level_number,
                    numbering_id=numbering_id,
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
                'numbering_id': block.numbering_id,
                'block_type': block.block_type,
                'index': block.index
            }
            blocks_data.append(block_data)
        
        # Create output data
        output_data = {
            'document_info': {
                'total_blocks': len(self.blocks),
                'content_blocks': len([b for b in self.blocks if b.block_type == "content"]),
                'blocks_with_levels': len([b for b in self.blocks if b.level_number is not None]),
                'blocks_without_levels': len([b for b in self.blocks if b.level_number is None])
            },
            'blocks': blocks_data
        }
        
        # Save to JSON
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"Content blocks saved to: {output_path}")
    
    def print_summary(self):
        """Print a simple summary of extracted content"""
        if not self.blocks:
            print("No content blocks found.")
            return
        
        content_blocks = [b for b in self.blocks if b.block_type == "content"]
        blocks_with_levels = [b for b in content_blocks if b.level_number is not None]
        
        print(f"=== CONTENT EXTRACTION SUMMARY ===")
        print(f"Total blocks: {len(self.blocks)}")
        print(f"Content blocks: {len(content_blocks)}")
        print(f"Blocks with existing levels: {len(blocks_with_levels)}")
        print(f"Blocks without levels: {len(content_blocks) - len(blocks_with_levels)}")
        
        # Show first few blocks
        print(f"\nFirst 10 blocks:")
        for i, block in enumerate(self.blocks[:10]):
            level_info = f" [Level {block.level_number}]" if block.level_number is not None else " [No level]"
            print(f"{i+1:2d}. [{block.block_type:15}] {block.text[:50]}{level_info}")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python simple_content_extractor.py <docx_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    extractor = SimpleContentExtractor()
    
    try:
        # Extract content blocks
        print(f"Extracting content blocks from: {docx_path}")
        blocks = extractor.extract_content_blocks(docx_path)
        print(f"Extracted {len(blocks)} content blocks")
        
        # Print summary
        extractor.print_summary()
        
        # Save to JSON in output directory
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_content_blocks.json")
        extractor.save_blocks_to_json(output_path)
        
        print(f"\nExtraction complete! Files saved to: {output_dir}")
        
    except Exception as e:
        print(f"Error extracting content blocks: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 