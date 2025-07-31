#!/usr/bin/env python3
"""
Multilist Level Formatting Analyzer

This script analyzes the JSON structure of Word documents to detect
and validate multilist level formatting patterns.
"""

import json
import sys
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from collections import defaultdict

@dataclass
class ListLevel:
    """Represents a list level with its properties"""
    text: str
    style_name: str
    numbering_id: Optional[int]
    numbering_level: Optional[int]
    index: int
    alignment: Optional[str] = None
    font_info: Optional[Dict[str, Any]] = None

@dataclass
class ListStructure:
    """Represents a complete list structure"""
    levels: List[ListLevel]
    numbering_ids: Dict[int, List[ListLevel]]
    style_patterns: Dict[str, List[ListLevel]]
    errors: List[str]
    warnings: List[str]

class MultilistAnalyzer:
    """Analyzes multilist level formatting in Word documents"""
    
    def __init__(self):
        self.structure = None
        self.analysis = {}
    
    def load_json_structure(self, json_path: str) -> Dict[str, Any]:
        """Load the JSON structure from file"""
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def extract_list_levels(self, data: Dict[str, Any]) -> List[ListLevel]:
        """Extract list levels from the document structure"""
        levels = []
        
        for paragraph in data.get('paragraphs', []):
            # Skip empty paragraphs
            if not paragraph.get('text', '').strip():
                continue
            
            # Extract numbering information
            numbering_info = paragraph.get('numbering', {})
            numbering_id = numbering_info.get('id')
            numbering_level = numbering_info.get('level')
            
            # Create ListLevel object
            level = ListLevel(
                text=paragraph.get('text', ''),
                style_name=paragraph.get('style_name', ''),
                numbering_id=numbering_id,
                numbering_level=numbering_level,
                index=paragraph.get('index', 0),
                alignment=paragraph.get('alignment'),
                font_info=self._extract_font_info(paragraph)
            )
            
            levels.append(level)
        
        return levels
    
    def _extract_font_info(self, paragraph: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Extract font information from paragraph runs"""
        runs = paragraph.get('runs', [])
        if not runs:
            return None
        
        # Get the first run's font info as representative
        first_run = runs[0]
        return {
            'font_name': first_run.get('font_name'),
            'font_size': first_run.get('font_size'),
            'bold': first_run.get('bold'),
            'italic': first_run.get('italic'),
            'underline': first_run.get('underline'),
            'font_color': first_run.get('font_color')
        }
    
    def analyze_list_structure(self, levels: List[ListLevel]) -> ListStructure:
        """Analyze the list structure for patterns and issues"""
        structure = ListStructure(
            levels=levels,
            numbering_ids=defaultdict(list),
            style_patterns=defaultdict(list),
            errors=[],
            warnings=[]
        )
        
        # Group by numbering ID
        for level in levels:
            if level.numbering_id is not None:
                structure.numbering_ids[level.numbering_id].append(level)
        
        # Group by style name
        for level in levels:
            if level.style_name:
                structure.style_patterns[level.style_name].append(level)
        
        # Analyze for issues
        self._detect_structure_issues(structure)
        
        return structure
    
    def _detect_structure_issues(self, structure: ListStructure):
        """Detect issues in the list structure"""
        
        # Check for missing numbering
        unnumbered_levels = [level for level in structure.levels if level.numbering_id is None]
        if unnumbered_levels:
            structure.warnings.append(f"Found {len(unnumbered_levels)} levels without numbering")
        
        # Check numbering ID consistency
        for numbering_id, levels in structure.numbering_ids.items():
            if len(levels) == 1:
                structure.warnings.append(f"Numbering ID {numbering_id} has only one level")
            
            # Check for level gaps
            levels_sorted = sorted(levels, key=lambda x: x.numbering_level or 0)
            for i in range(len(levels_sorted) - 1):
                current_level = levels_sorted[i].numbering_level or 0
                next_level = levels_sorted[i + 1].numbering_level or 0
                if next_level - current_level > 1:
                    structure.warnings.append(
                        f"Gap in numbering levels: {current_level} -> {next_level} "
                        f"in numbering ID {numbering_id}"
                    )
        
        # Check style consistency
        for style_name, levels in structure.style_patterns.items():
            if len(levels) == 1:
                structure.warnings.append(f"Style '{style_name}' used only once")
            
            # Check if style has consistent numbering
            numbering_ids = set(level.numbering_id for level in levels if level.numbering_id is not None)
            if len(numbering_ids) > 1:
                structure.warnings.append(
                    f"Style '{style_name}' used across multiple numbering IDs: {numbering_ids}"
                )
    
    def generate_analysis_report(self, structure: ListStructure) -> Dict[str, Any]:
        """Generate a comprehensive analysis report"""
        report = {
            'summary': {
                'total_levels': len(structure.levels),
                'numbered_levels': len([l for l in structure.levels if l.numbering_id is not None]),
                'unnumbered_levels': len([l for l in structure.levels if l.numbering_id is None]),
                'unique_numbering_ids': len(structure.numbering_ids),
                'unique_styles': len(structure.style_patterns),
                'errors': len(structure.errors),
                'warnings': len(structure.warnings)
            },
            'numbering_analysis': {},
            'style_analysis': {},
            'errors': structure.errors,
            'warnings': structure.warnings,
            'recommendations': []
        }
        
        # Analyze numbering patterns
        for numbering_id, levels in structure.numbering_ids.items():
            report['numbering_analysis'][f'numbering_id_{numbering_id}'] = {
                'level_count': len(levels),
                'levels': [l.numbering_level for l in levels if l.numbering_level is not None],
                'styles': list(set(l.style_name for l in levels)),
                'text_samples': [l.text[:50] for l in levels[:3]]  # First 3 texts
            }
        
        # Analyze style patterns
        for style_name, levels in structure.style_patterns.items():
            report['style_analysis'][style_name] = {
                'usage_count': len(levels),
                'numbering_ids': list(set(l.numbering_id for l in levels if l.numbering_id is not None)),
                'text_samples': [l.text[:50] for l in levels[:3]]  # First 3 texts
            }
        
        # Generate recommendations
        if structure.warnings:
            report['recommendations'].append("Review warnings for potential formatting issues")
        
        if len(structure.numbering_ids) > 1:
            report['recommendations'].append("Consider consolidating multiple numbering schemes")
        
        unnumbered_count = len([l for l in structure.levels if l.numbering_id is None])
        if unnumbered_count > len(structure.levels) * 0.5:
            report['recommendations'].append("High percentage of unnumbered content - consider adding numbering")
        
        return report
    
    def analyze_document(self, json_path: str) -> Dict[str, Any]:
        """Complete analysis of a document"""
        print(f"Analyzing document: {json_path}")
        
        # Load JSON structure
        data = self.load_json_structure(json_path)
        
        # Extract list levels
        levels = self.extract_list_levels(data)
        print(f"Extracted {len(levels)} list levels")
        
        # Analyze structure
        structure = self.analyze_list_structure(levels)
        
        # Generate report
        report = self.generate_analysis_report(structure)
        
        return {
            'document_path': json_path,
            'structure': structure,
            'analysis': report
        }
    
    def save_analysis_report(self, analysis: Dict[str, Any], output_path: str):
        """Save the analysis report to JSON"""
        # Convert dataclasses to dictionaries for JSON serialization
        report_data = {
            'document_path': analysis['document_path'],
            'analysis': analysis['analysis'],
            'structure_summary': {
                'total_levels': len(analysis['structure'].levels),
                'numbering_ids': list(analysis['structure'].numbering_ids.keys()),
                'styles': list(analysis['structure'].style_patterns.keys()),
                'errors': analysis['structure'].errors,
                'warnings': analysis['structure'].warnings
            }
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"Analysis report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python multilist_analyzer.py <json_file> [output_file]")
        sys.exit(1)
    
    json_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(json_path):
        print(f"Error: File not found: {json_path}")
        sys.exit(1)
    
    analyzer = MultilistAnalyzer()
    
    try:
        # Analyze the document
        analysis = analyzer.analyze_document(json_path)
        
        # Print summary
        summary = analysis['analysis']['summary']
        print(f"\nAnalysis Summary:")
        print(f"  Total levels: {summary['total_levels']}")
        print(f"  Numbered levels: {summary['numbered_levels']}")
        print(f"  Unnumbered levels: {summary['unnumbered_levels']}")
        print(f"  Unique numbering IDs: {summary['unique_numbering_ids']}")
        print(f"  Unique styles: {summary['unique_styles']}")
        print(f"  Errors: {summary['errors']}")
        print(f"  Warnings: {summary['warnings']}")
        
        # Print warnings and errors
        if analysis['structure'].warnings:
            print(f"\nWarnings:")
            for warning in analysis['structure'].warnings:
                print(f"  - {warning}")
        
        if analysis['structure'].errors:
            print(f"\nErrors:")
            for error in analysis['structure'].errors:
                print(f"  - {error}")
        
        # Save report if output path specified
        if output_path:
            analyzer.save_analysis_report(analysis, output_path)
        else:
            # Save with default name
            base_name = Path(json_path).stem
            default_output = f"{base_name}_analysis.json"
            analyzer.save_analysis_report(analysis, default_output)
        
    except Exception as e:
        print(f"Error analyzing document: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 