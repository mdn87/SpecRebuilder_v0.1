#!/usr/bin/env python3
"""
Numbering Analyzer

This script analyzes the numbering information in Word documents to find
the relationship between text content and actual numbering values.
"""

import json
import sys
import os
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from word_to_json import WordToJsonConverter

@dataclass
class NumberingInfo:
    """Represents numbering information for a paragraph"""
    text: str
    numbering_id: Optional[int]
    level: Optional[int]
    style_name: str
    index: int
    actual_number: Optional[str] = None

class NumberingAnalyzer:
    """Analyzes numbering information in Word documents"""
    
    def __init__(self):
        self.numbering_data = {}
        self.paragraphs = []
    
    def extract_numbering_from_docx(self, docx_path: str) -> Dict[str, Any]:
        """Extract numbering information directly from the docx file"""
        import zipfile
        from xml.etree import ElementTree as ET
        
        numbering_info = {}
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                # Extract numbering.xml
                if 'word/numbering.xml' in zip_file.namelist():
                    numbering_xml = zip_file.read('word/numbering.xml')
                    root = ET.fromstring(numbering_xml)
                    
                    # Parse numbering definitions
                    for num in root.findall('.//w:num', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        num_id = num.get('w:numId')
                        if num_id:
                            numbering_info[f'num_{num_id}'] = {
                                'id': num_id,
                                'abstract_num_id': None,
                                'levels': {}
                            }
                            
                            # Get abstract numbering reference
                            abstract_num_ref = num.find('.//w:abstractNumId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                            if abstract_num_ref is not None:
                                numbering_info[f'num_{num_id}']['abstract_num_id'] = abstract_num_ref.get('w:val')
                    
                    # Parse abstract numbering definitions
                    for abstract_num in root.findall('.//w:abstractNum', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        abstract_num_id = abstract_num.get('w:abstractNumId')
                        if abstract_num_id:
                            numbering_info[f'abstract_{abstract_num_id}'] = {
                                'id': abstract_num_id,
                                'levels': {}
                            }
                            
                            # Parse level definitions
                            for level in abstract_num.findall('.//w:lvl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                                level_id = level.get('w:ilvl')
                                if level_id is not None:
                                    level_info = {
                                        'id': level_id,
                                        'start': None,
                                        'num_fmt': None,
                                        'lvl_text': None,
                                        'lvl_jc': None
                                    }
                                    
                                    # Get start value
                                    start = level.find('.//w:start', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                                    if start is not None:
                                        level_info['start'] = start.get('w:val')
                                    
                                    # Get number format
                                    num_fmt = level.find('.//w:numFmt', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                                    if num_fmt is not None:
                                        level_info['num_fmt'] = num_fmt.get('w:val')
                                    
                                    # Get level text
                                    lvl_text = level.find('.//w:lvlText', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                                    if lvl_text is not None:
                                        level_info['lvl_text'] = lvl_text.get('w:val')
                                    
                                    # Get justification
                                    lvl_jc = level.find('.//w:lvlJc', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                                    if lvl_jc is not None:
                                        level_info['lvl_jc'] = lvl_jc.get('w:val')
                                    
                                    numbering_info[f'abstract_{abstract_num_id}']['levels'][level_id] = level_info
                
                # Extract document.xml to get paragraph numbering references
                if 'word/document.xml' in zip_file.namelist():
                    doc_xml = zip_file.read('word/document.xml')
                    doc_root = ET.fromstring(doc_xml)
                    
                    # Find paragraphs with numbering
                    for i, p in enumerate(doc_root.findall('.//w:p', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})):
                        ppr = p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                        if ppr is not None:
                            num_pr = ppr.find('.//w:numPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                            if num_pr is not None:
                                num_id = num_pr.find('.//w:numId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                                ilvl = num_pr.find('.//w:ilvl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                                
                                if num_id is not None and ilvl is not None:
                                    numbering_info[f'paragraph_{i}'] = {
                                        'num_id': num_id.get('w:val'),
                                        'level': ilvl.get('w:val'),
                                        'index': i
                                    }
        
        except Exception as e:
            print(f"Error extracting numbering from docx: {e}")
        
        return numbering_info
    
    def analyze_numbering_relationships(self, docx_path: str) -> Dict[str, Any]:
        """Analyze the relationship between text and numbering"""
        # First get the basic structure
        converter = WordToJsonConverter()
        json_path = converter.convert_to_json(docx_path)
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Extract numbering information from docx
        numbering_info = self.extract_numbering_from_docx(docx_path)
        
        # Analyze paragraphs with numbering
        numbered_paragraphs = []
        for paragraph in data.get('paragraphs', []):
            numbering = paragraph.get('numbering', {})
            if numbering:
                numbered_paragraphs.append({
                    'text': paragraph.get('text', ''),
                    'numbering_id': numbering.get('id'),
                    'level': numbering.get('level'),
                    'style_name': paragraph.get('style_name', ''),
                    'index': paragraph.get('index', 0)
                })
        
        # Try to find BWA-SUBSECTION1 specifically
        bwa_subsections = []
        for paragraph in data.get('paragraphs', []):
            if 'BWA-SUBSECTION' in paragraph.get('text', ''):
                bwa_subsections.append({
                    'text': paragraph.get('text', ''),
                    'numbering': paragraph.get('numbering', {}),
                    'style_name': paragraph.get('style_name', ''),
                    'index': paragraph.get('index', 0)
                })
        
        return {
            'numbering_definitions': numbering_info,
            'numbered_paragraphs': numbered_paragraphs,
            'bwa_subsections': bwa_subsections,
            'total_paragraphs': len(data.get('paragraphs', [])),
            'paragraphs_with_numbering': len(numbered_paragraphs)
        }
    
    def print_analysis(self, analysis: Dict[str, Any]):
        """Print the numbering analysis"""
        print("=== NUMBERING ANALYSIS ===")
        print(f"Total paragraphs: {analysis['total_paragraphs']}")
        print(f"Paragraphs with numbering: {analysis['paragraphs_with_numbering']}")
        print()
        
        print("Numbering Definitions:")
        for key, value in analysis['numbering_definitions'].items():
            if key.startswith('num_') or key.startswith('abstract_'):
                print(f"  {key}: {value}")
        print()
        
        print("BWA Subsections Found:")
        for subsection in analysis['bwa_subsections']:
            print(f"  Text: {subsection['text']}")
            print(f"  Style: {subsection['style_name']}")
            print(f"  Numbering: {subsection['numbering']}")
            print(f"  Index: {subsection['index']}")
            print()
        
        print("All Numbered Paragraphs:")
        for para in analysis['numbered_paragraphs']:
            print(f"  [{para['index']:2d}] {para['text'][:50]:<50} | ID:{para['numbering_id']} Level:{para['level']} | Style:{para['style_name']}")
    
    def save_analysis(self, analysis: Dict[str, Any], output_path: str):
        """Save the analysis to JSON"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(analysis, f, indent=2, ensure_ascii=False, default=str)
        print(f"Analysis saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python numbering_analyzer.py <docx_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    analyzer = NumberingAnalyzer()
    
    try:
        print(f"Analyzing numbering in: {docx_path}")
        analysis = analyzer.analyze_numbering_relationships(docx_path)
        
        # Print analysis
        analyzer.print_analysis(analysis)
        
        # Save analysis
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_numbering_analysis.json")
        analyzer.save_analysis(analysis, output_path)
        
        print(f"\nAnalysis complete! Files saved to: {output_dir}")
        
    except Exception as e:
        print(f"Error analyzing numbering: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 