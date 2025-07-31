#!/usr/bin/env python3
"""
Complete Content Block Analysis Pipeline

This script combines content block extraction and pattern analysis
to provide a complete analysis of Word document structure.
"""

import sys
import os
from pathlib import Path
from typing import Dict, Any
from content_block_extractor import ContentBlockExtractor
from block_pattern_analyzer import BlockPatternAnalyzer

def analyze_document_complete(docx_path: str, output_dir: str = None) -> Dict[str, Any]:
    """Complete analysis pipeline for a Word document"""
    
    # Create output directory if not specified
    if output_dir is None:
        output_dir = "."
    
    print(f"=== COMPLETE CONTENT BLOCK ANALYSIS ===")
    print(f"Document: {docx_path}")
    print()
    
    # Step 1: Extract content blocks
    print("Step 1: Extracting content blocks...")
    extractor = ContentBlockExtractor()
    blocks = extractor.extract_content_blocks(docx_path)
    print(f"Extracted {len(blocks)} content blocks")
    
    # Save content blocks
    base_name = Path(docx_path).stem
    content_blocks_path = os.path.join(output_dir, f"{base_name}_content_blocks.json")
    extractor.save_blocks_to_json(content_blocks_path)
    
    # Step 2: Analyze patterns and suggest levels
    print("\nStep 2: Analyzing patterns and suggesting levels...")
    analyzer = BlockPatternAnalyzer()
    analyzer.load_blocks_from_json(content_blocks_path)
    
    # Generate pattern analysis
    pattern_report = analyzer.generate_analysis_report()
    print(pattern_report)
    
    # Save suggestions
    suggestions_path = os.path.join(output_dir, f"{base_name}_suggestions.json")
    analyzer.save_suggestions_to_json(suggestions_path)
    
    # Step 3: Generate summary
    print("\nStep 3: Generating summary...")
    analysis = extractor.analyze_level_distribution()
    suggestions = analyzer.suggest_levels_for_missing_blocks()
    
    summary = {
        'document_path': docx_path,
        'total_blocks': len(blocks),
        'content_blocks': len([b for b in blocks if b.block_type == "content"]),
        'blocks_with_levels': len([b for b in blocks if b.block_type == "content" and b.level_number is not None]),
        'blocks_without_levels': len([b for b in blocks if b.block_type == "content" and b.level_number is None]),
        'suggestions_made': len([s for s in suggestions if s['suggested_level'] is not None]),
        'files_generated': {
            'content_blocks': content_blocks_path,
            'suggestions': suggestions_path
        }
    }
    
    # Print summary
    print("\n=== ANALYSIS SUMMARY ===")
    print(f"Document: {docx_path}")
    print(f"Total blocks: {summary['total_blocks']}")
    print(f"Content blocks: {summary['content_blocks']}")
    print(f"Blocks with levels: {summary['blocks_with_levels']}")
    print(f"Blocks without levels: {summary['blocks_without_levels']}")
    print(f"Level suggestions made: {summary['suggestions_made']}")
    print()
    print("Files generated:")
    print(f"  Content blocks: {content_blocks_path}")
    print(f"  Suggestions: {suggestions_path}")
    
    return summary

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python complete_analysis.py <docx_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    try:
        result = analyze_document_complete(docx_path, output_dir)
        print(f"\nAnalysis complete!")
        
    except Exception as e:
        print(f"Error analyzing document: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 