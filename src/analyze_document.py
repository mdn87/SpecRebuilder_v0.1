#!/usr/bin/env python3
"""
Document Analysis Pipeline

This script combines Word to JSON conversion and multilist analysis
in a single pipeline for easy document analysis.
"""

import sys
import os
from pathlib import Path
from typing import Dict, Any
from word_to_json import WordToJsonConverter
from multilist_analyzer import MultilistAnalyzer

def analyze_word_document(docx_path: str, output_dir: str = None) -> Dict[str, Any]:
    """Complete analysis pipeline for a Word document"""
    
    # Create output directory if not specified
    if output_dir is None:
        output_dir = "."
    
    # Step 1: Convert Word to JSON
    print(f"Step 1: Converting {docx_path} to JSON...")
    converter = WordToJsonConverter()
    json_path = converter.convert_to_json(docx_path)
    
    # Step 2: Analyze the JSON structure
    print(f"Step 2: Analyzing multilist structure...")
    analyzer = MultilistAnalyzer()
    analysis = analyzer.analyze_document(json_path)
    
    # Step 3: Save analysis report
    base_name = Path(docx_path).stem
    analysis_path = os.path.join(output_dir, f"{base_name}_analysis.json")
    analyzer.save_analysis_report(analysis, analysis_path)
    
    # Print summary
    summary = analysis['analysis']['summary']
    print(f"\n=== ANALYSIS SUMMARY ===")
    print(f"Document: {docx_path}")
    print(f"Total levels: {summary['total_levels']}")
    print(f"Numbered levels: {summary['numbered_levels']}")
    print(f"Unnumbered levels: {summary['unnumbered_levels']}")
    print(f"Unique numbering IDs: {summary['unique_numbering_ids']}")
    print(f"Unique styles: {summary['unique_styles']}")
    print(f"Errors: {summary['errors']}")
    print(f"Warnings: {summary['warnings']}")
    
    # Print warnings
    if analysis['structure'].warnings:
        print(f"\n=== WARNINGS ===")
        for warning in analysis['structure'].warnings:
            print(f"  ⚠️  {warning}")
    
    # Print recommendations
    if analysis['analysis']['recommendations']:
        print(f"\n=== RECOMMENDATIONS ===")
        for rec in analysis['analysis']['recommendations']:
            print(f"  💡 {rec}")
    
    # Print numbering analysis
    if analysis['analysis']['numbering_analysis']:
        print(f"\n=== NUMBERING ANALYSIS ===")
        for numbering_id, info in analysis['analysis']['numbering_analysis'].items():
            print(f"  {numbering_id}: {info['level_count']} levels, styles: {info['styles']}")
    
    # Print style analysis
    if analysis['analysis']['style_analysis']:
        print(f"\n=== STYLE ANALYSIS ===")
        for style_name, info in analysis['analysis']['style_analysis'].items():
            numbering_info = f" (numbering IDs: {info['numbering_ids']})" if info['numbering_ids'] else " (no numbering)"
            print(f"  {style_name}: {info['usage_count']} uses{numbering_info}")
    
    return {
        'json_path': json_path,
        'analysis_path': analysis_path,
        'analysis': analysis
    }

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python analyze_document.py <docx_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    try:
        result = analyze_word_document(docx_path, output_dir)
        print(f"\n=== FILES GENERATED ===")
        print(f"JSON structure: {result['json_path']}")
        print(f"Analysis report: {result['analysis_path']}")
        
    except Exception as e:
        print(f"Error analyzing document: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 