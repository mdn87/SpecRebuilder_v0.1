#!/usr/bin/env python3
"""
Comprehensive Numbering Analyzer

This script analyzes all possible locations where numbering data might be stored
in a Word document and compares it to expected numbering from a text file.
"""

import json
import sys
import os
import re
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from word_to_json import WordToJsonConverter

@dataclass
class ExpectedNumbering:
    """Represents expected numbering from text file"""
    line_number: int
    expected_number: str
    content: str
    level: Optional[int] = None

@dataclass
class FoundNumbering:
    """Represents numbering found in Word document"""
    location: str
    numbering_data: Dict[str, Any]
    confidence: float
    description: str

class ComprehensiveNumberingAnalyzer:
    """Analyzes all possible numbering data locations"""
    
    def __init__(self):
        self.expected_numbering = []
        self.found_numbering = []
        self.analysis_results = {}
    
    def load_expected_numbering(self, txt_path: str) -> List[ExpectedNumbering]:
        """Load expected numbering from text file"""
        expected = []
        
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for i, line in enumerate(lines, 1):
            line = line.strip()
            if not line or line in ['SECTION 00 00 00', 'SECTION TITLE', 'END OF SECTION']:
                continue
            
            # Parse line to extract numbering and content
            parts = line.split('\t')
            if len(parts) >= 2:
                expected_number = parts[0].strip()
                content = parts[1].strip()
                
                # Determine level based on numbering pattern
                level = self.determine_level_from_numbering(expected_number)
                
                expected.append(ExpectedNumbering(
                    line_number=i,
                    expected_number=expected_number,
                    content=content,
                    level=level
                ))
        
        self.expected_numbering = expected
        return expected
    
    def determine_level_from_numbering(self, numbering: str) -> Optional[int]:
        """Determine level based on numbering pattern"""
        if re.match(r'^\d+\.0$', numbering):  # 1.0, 2.0
            return 0
        elif re.match(r'^\d+\.\d+$', numbering):  # 1.01, 1.02
            return 1
        elif re.match(r'^[A-Z]\.$', numbering):  # A., B.
            return 2
        elif re.match(r'^\d+\.$', numbering):  # 1., 2.
            return 3
        elif re.match(r'^[a-z]\.$', numbering):  # a., b.
            return 4
        elif re.match(r'^[ivxlcdm]+\.$', numbering, re.IGNORECASE):  # i., ii.
            return 5
        return None
    
    def extract_all_numbering_locations(self, docx_path: str) -> Dict[str, Any]:
        """Extract numbering data from all possible locations"""
        import zipfile
        from xml.etree import ElementTree as ET
        
        numbering_locations = {
            'numbering_xml': {},
            'document_xml': {},
            'styles_xml': {},
            'paragraph_properties': {},
            'runs_with_numbering': {},
            'content_with_numbering': {}
        }
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                # 1. Extract numbering.xml
                if 'word/numbering.xml' in zip_file.namelist():
                    numbering_xml = zip_file.read('word/numbering.xml')
                    root = ET.fromstring(numbering_xml)
                    numbering_locations['numbering_xml'] = self.parse_numbering_xml(root)
                
                # 2. Extract document.xml
                if 'word/document.xml' in zip_file.namelist():
                    doc_xml = zip_file.read('word/document.xml')
                    doc_root = ET.fromstring(doc_xml)
                    numbering_locations['document_xml'] = self.parse_document_xml(doc_root)
                
                # 3. Extract styles.xml
                if 'word/styles.xml' in zip_file.namelist():
                    styles_xml = zip_file.read('word/styles.xml')
                    styles_root = ET.fromstring(styles_xml)
                    numbering_locations['styles_xml'] = self.parse_styles_xml(styles_root)
        
        except Exception as e:
            print(f"Error extracting numbering locations: {e}")
        
        return numbering_locations
    
    def parse_numbering_xml(self, root) -> Dict[str, Any]:
        """Parse numbering.xml for numbering definitions"""
        result = {
            'abstract_nums': {},
            'nums': {},
            'levels': {}
        }
        
        # Parse abstract numbering definitions
        for abstract_num in root.findall('.//w:abstractNum', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            abstract_num_id = abstract_num.get('w:abstractNumId')
            if abstract_num_id:
                result['abstract_nums'][abstract_num_id] = {
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
                        
                        result['abstract_nums'][abstract_num_id]['levels'][level_id] = level_info
        
        # Parse numbering instances
        for num in root.findall('.//w:num', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            num_id = num.get('w:numId')
            if num_id:
                result['nums'][num_id] = {
                    'id': num_id,
                    'abstract_num_id': None
                }
                
                # Get abstract numbering reference
                abstract_num_ref = num.find('.//w:abstractNumId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if abstract_num_ref is not None:
                    result['nums'][num_id]['abstract_num_id'] = abstract_num_ref.get('w:val')
        
        return result
    
    def parse_document_xml(self, root) -> Dict[str, Any]:
        """Parse document.xml for paragraph numbering references"""
        result = {
            'paragraphs_with_numbering': [],
            'runs_with_numbering': [],
            'content_with_numbering': []
        }
        
        # Find paragraphs with numbering
        for i, p in enumerate(root.findall('.//w:p', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})):
            ppr = p.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if ppr is not None:
                num_pr = ppr.find('.//w:numPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if num_pr is not None:
                    num_id = num_pr.find('.//w:numId', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    ilvl = num_pr.find('.//w:ilvl', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    
                    if num_id is not None and ilvl is not None:
                        result['paragraphs_with_numbering'].append({
                            'index': i,
                            'num_id': num_id.get('w:val'),
                            'level': ilvl.get('w:val')
                        })
        
        return result
    
    def parse_styles_xml(self, root) -> Dict[str, Any]:
        """Parse styles.xml for numbering-related styles"""
        result = {
            'styles_with_numbering': [],
            'numbering_styles': {}
        }
        
        # Find styles that might be related to numbering
        for style in root.findall('.//w:style', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
            style_id = style.get('w:styleId')
            if style_id:
                # Check if style has numbering properties
                ppr = style.find('.//w:pPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if ppr is not None:
                    num_pr = ppr.find('.//w:numPr', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if num_pr is not None:
                        result['styles_with_numbering'].append({
                            'style_id': style_id,
                            'num_pr': self.element_to_dict(num_pr)
                        })
        
        return result
    
    def element_to_dict(self, element) -> Dict[str, Any]:
        """Convert XML element to dictionary"""
        result = {}
        for child in element:
            tag = child.tag.split('}')[-1]  # Remove namespace
            if len(child) == 0:
                result[tag] = child.get('w:val') if child.get('w:val') else child.text
            else:
                result[tag] = self.element_to_dict(child)
        return result
    
    def analyze_word_document_structure(self, docx_path: str) -> Dict[str, Any]:
        """Analyze the Word document structure for numbering"""
        # Get basic structure
        converter = WordToJsonConverter()
        json_path = converter.convert_to_json(docx_path)
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Extract all numbering locations
        numbering_locations = self.extract_all_numbering_locations(docx_path)
        
        # Analyze paragraphs for numbering
        paragraphs_with_numbering = []
        for paragraph in data.get('paragraphs', []):
            numbering = paragraph.get('numbering', {})
            if numbering:
                paragraphs_with_numbering.append({
                    'text': paragraph.get('text', ''),
                    'numbering_id': numbering.get('id'),
                    'level': numbering.get('level'),
                    'style_name': paragraph.get('style_name', ''),
                    'index': paragraph.get('index', 0)
                })
        
        return {
            'numbering_locations': numbering_locations,
            'paragraphs_with_numbering': paragraphs_with_numbering,
            'total_paragraphs': len(data.get('paragraphs', [])),
            'paragraphs_with_numbering_count': len(paragraphs_with_numbering)
        }
    
    def generate_comprehensive_report(self, docx_path: str, txt_path: str) -> Dict[str, Any]:
        """Generate a comprehensive numbering analysis report"""
        
        # Load expected numbering
        expected = self.load_expected_numbering(txt_path)
        
        # Analyze Word document
        word_analysis = self.analyze_word_document_structure(docx_path)
        
        # Create possible numbering locations list
        possible_locations = [
            "Word numbering.xml definitions",
            "Document.xml paragraph properties",
            "Styles.xml numbering properties", 
            "Paragraph runs with numbering",
            "Content blocks with numbering",
            "Numbering found within content block as plain text",
            "No numbering data available"
        ]
        
        # Analyze each expected numbering against found data
        numbering_analysis = []
        for expected_item in expected:
            analysis = {
                'expected': {
                    'line_number': expected_item.line_number,
                    'numbering': expected_item.expected_number,
                    'content': expected_item.content,
                    'level': expected_item.level
                },
                'found_locations': [],
                'best_match': None,
                'confidence': 0.0
            }
            
            # Check each possible location
            for location in possible_locations:
                match = self.check_location_for_numbering(
                    location, expected_item, word_analysis
                )
                if match:
                    analysis['found_locations'].append({
                        'location': location,
                        'data': match,
                        'confidence': match.get('confidence', 0.0)
                    })
            
            # Find best match
            if analysis['found_locations']:
                best_match = max(analysis['found_locations'], key=lambda x: x['confidence'])
                analysis['best_match'] = best_match
                analysis['confidence'] = best_match['confidence']
            
            numbering_analysis.append(analysis)
        
        return {
            'expected_numbering_count': len(expected),
            'word_analysis': word_analysis,
            'numbering_analysis': numbering_analysis,
            'possible_locations': possible_locations,
            'summary': self.generate_summary(expected, word_analysis, numbering_analysis)
        }
    
    def check_location_for_numbering(self, location: str, expected: ExpectedNumbering, word_analysis: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Check a specific location for numbering data"""
        
        if location == "Word numbering.xml definitions":
            # Check if numbering.xml has definitions that could produce this numbering
            numbering_defs = word_analysis['numbering_locations'].get('numbering_xml', {})
            if numbering_defs.get('abstract_nums'):
                return {
                    'data': numbering_defs,
                    'confidence': 0.8,
                    'description': 'Numbering definitions found in numbering.xml'
                }
        
        elif location == "Document.xml paragraph properties":
            # Check if document.xml has paragraph numbering properties
            doc_xml = word_analysis['numbering_locations'].get('document_xml', {})
            if doc_xml.get('paragraphs_with_numbering'):
                return {
                    'data': doc_xml,
                    'confidence': 0.7,
                    'description': 'Paragraph numbering properties found in document.xml'
                }
        
        elif location == "Styles.xml numbering properties":
            # Check if styles.xml has numbering properties
            styles_xml = word_analysis['numbering_locations'].get('styles_xml', {})
            if styles_xml.get('styles_with_numbering'):
                return {
                    'data': styles_xml,
                    'confidence': 0.6,
                    'description': 'Numbering properties found in styles.xml'
                }
        
        elif location == "Paragraph runs with numbering":
            # Check if any paragraphs have numbering applied
            paragraphs = word_analysis.get('paragraphs_with_numbering', [])
            for para in paragraphs:
                if expected.content in para.get('text', ''):
                    return {
                        'data': para,
                        'confidence': 0.9,
                        'description': f"Numbering found in paragraph: {para.get('text', '')[:50]}"
                    }
        
        elif location == "Content blocks with numbering":
            # Check if content has numbering in the text itself
            if expected.expected_number in expected.content or expected.expected_number in f"{expected.expected_number}\t{expected.content}":
                return {
                    'data': {'text': expected.content, 'numbering': expected.expected_number},
                    'confidence': 0.5,
                    'description': 'Numbering appears to be part of content text'
                }
        
        elif location == "Numbering found within content block as plain text":
            # This is a fallback - check if numbering appears anywhere in the content
            paragraphs = word_analysis.get('paragraphs_with_numbering', [])
            for para in paragraphs:
                if expected.expected_number in para.get('text', ''):
                    return {
                        'data': para,
                        'confidence': 0.4,
                        'description': f"Numbering '{expected.expected_number}' found in text: {para.get('text', '')[:50]}"
                    }
        
        elif location == "No numbering data available":
            # This is the final fallback
            return {
                'data': None,
                'confidence': 0.0,
                'description': 'No numbering data found for this item'
            }
        
        return None
    
    def generate_summary(self, expected: List[ExpectedNumbering], word_analysis: Dict[str, Any], numbering_analysis: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate a summary of the analysis"""
        
        total_expected = len(expected)
        found_with_numbering = len([a for a in numbering_analysis if a['best_match'] and a['confidence'] > 0.5])
        found_without_numbering = len([a for a in numbering_analysis if not a['best_match'] or a['confidence'] <= 0.5])
        
        location_counts = {}
        for analysis in numbering_analysis:
            if analysis['best_match']:
                location = analysis['best_match']['location']
                location_counts[location] = location_counts.get(location, 0) + 1
        
        return {
            'total_expected_items': total_expected,
            'found_with_numbering': found_with_numbering,
            'found_without_numbering': found_without_numbering,
            'location_distribution': location_counts,
            'overall_confidence': sum(a['confidence'] for a in numbering_analysis) / len(numbering_analysis) if numbering_analysis else 0.0
        }
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save the comprehensive report"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"Comprehensive report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python comprehensive_numbering_analyzer.py <docx_file> <txt_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    txt_path = sys.argv[2]
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "output"
    
    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found: {docx_path}")
        sys.exit(1)
    
    if not os.path.exists(txt_path):
        print(f"Error: TXT file not found: {txt_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    analyzer = ComprehensiveNumberingAnalyzer()
    
    try:
        print(f"Analyzing numbering in: {docx_path}")
        print(f"Comparing with expected numbering in: {txt_path}")
        
        # Generate comprehensive report
        report = analyzer.generate_comprehensive_report(docx_path, txt_path)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_comprehensive_numbering_report.json")
        analyzer.save_report(report, output_path)
        
        # Print summary
        summary = report['summary']
        print(f"\n=== ANALYSIS SUMMARY ===")
        print(f"Total expected items: {summary['total_expected_items']}")
        print(f"Found with numbering: {summary['found_with_numbering']}")
        print(f"Found without numbering: {summary['found_without_numbering']}")
        print(f"Overall confidence: {summary['overall_confidence']:.2f}")
        print(f"\nLocation distribution:")
        for location, count in summary['location_distribution'].items():
            print(f"  {location}: {count}")
        
        print(f"\nComprehensive report saved to: {output_path}")
        
    except Exception as e:
        print(f"Error analyzing numbering: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 