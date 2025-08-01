#!/usr/bin/env python3
"""
Text Preview Generator

This script takes JSON analysis data and generates a text preview
of what the reconstructed document would look like.
"""

import os
import sys
import json
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

@dataclass
class ParagraphData:
    """Represents a paragraph with numbering information"""
    index: int
    list_number: str
    text: str
    level: Optional[int] = None
    inferred_number: Optional[str] = None
    cleaned_content: Optional[str] = None

class TextPreviewGenerator:
    """Generates text preview from JSON analysis data"""
    
    def load_json_analysis(self, json_path: str) -> Dict[str, Any]:
        """Load the JSON analysis data"""
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def parse_paragraphs_from_json(self, json_data: Dict[str, Any]) -> List[ParagraphData]:
        """Parse paragraphs from JSON data"""
        paragraphs = []
        
        for para_data in json_data.get('all_paragraphs', []):
            paragraph = ParagraphData(
                index=para_data.get('index', 0),
                list_number=para_data.get('list_number', ''),
                text=para_data.get('text', ''),
                level=para_data.get('level'),
                inferred_number=para_data.get('inferred_number'),
                cleaned_content=para_data.get('cleaned_content')
            )
            paragraphs.append(paragraph)
        
        return paragraphs
    
    def format_paragraph_text(self, para_data: ParagraphData) -> str:
        """Format paragraph text with appropriate numbering and indentation"""
        # Skip empty paragraphs
        if not para_data.text.strip():
            return ""
        
        # Determine if this should be numbered
        has_numbering = bool(para_data.list_number or para_data.inferred_number)
        
        if has_numbering:
            # Get the numbering to use
            numbering = para_data.list_number if para_data.list_number else para_data.inferred_number
            level = para_data.level if para_data.level is not None else 0
            
            # Get the content
            if para_data.cleaned_content:
                content = para_data.cleaned_content
            else:
                content = para_data.text
            
            # Add indentation based on level
            indent = "  " * level
            
            # Format: indent + numbering + space + content
            return f"{indent}{numbering} {content}"
        
        else:
            # Regular paragraph (no numbering)
            return para_data.text
    
    def generate_preview(self, json_path: str, output_path: str):
        """Generate text preview from JSON analysis"""
        print(f"Loading JSON analysis from: {json_path}")
        json_data = self.load_json_analysis(json_path)
        
        print(f"Parsing {len(json_data.get('all_paragraphs', []))} paragraphs...")
        paragraphs = self.parse_paragraphs_from_json(json_data)
        
        print(f"Generating text preview with {len(paragraphs)} paragraphs...")
        
        # Generate preview text
        preview_lines = []
        for para_data in paragraphs:
            formatted_text = self.format_paragraph_text(para_data)
            if formatted_text:
                preview_lines.append(formatted_text)
        
        # Write to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(preview_lines))
        
        print(f"Text preview saved to: {output_path}")
        
        # Also print first 20 lines to console
        print("\n=== PREVIEW (First 20 lines) ===")
        for i, line in enumerate(preview_lines[:20]):
            print(f"{i+1:2d}: {line}")
        
        if len(preview_lines) > 20:
            print(f"... and {len(preview_lines) - 20} more lines")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python text_preview_generator.py <json_file> <output_txt>")
        print("Example: python text_preview_generator.py output/SECTION_00_00_00_hybrid_analysis.json preview_SECTION_00_00_00.txt")
        sys.exit(1)
    
    json_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(json_path):
        print(f"Error: JSON file not found: {json_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    try:
        generator = TextPreviewGenerator()
        generator.generate_preview(json_path, output_path)
        
    except Exception as e:
        print(f"Error in text preview generation: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 