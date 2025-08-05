#!/usr/bin/env python3
"""
Word Document Structure Analyzer

This script analyzes the internal structure of Word documents
to identify differences and potential issues.
"""

import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Any

class DocxAnalyzer:
    """Analyzes Word document structure"""
    
    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
    
    def analyze_document(self, docx_path: str) -> Dict[str, Any]:
        """Analyze a Word document and return its structure"""
        if not os.path.exists(docx_path):
            print(f"Error: Document not found: {docx_path}")
            return {}
        
        analysis = {
            'file_size': os.path.getsize(docx_path),
            'files': [],
            'content_types': {},
            'relationships': {},
            'document_structure': {},
            'numbering_structure': {}
        }
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as zipf:
                # List all files
                analysis['files'] = zipf.namelist()
                
                # Analyze content types
                if '[Content_Types].xml' in zipf.namelist():
                    content_types_xml = zipf.read('[Content_Types].xml').decode('utf-8')
                    analysis['content_types'] = self.parse_content_types(content_types_xml)
                
                # Analyze relationships
                if '_rels/.rels' in zipf.namelist():
                    rels_xml = zipf.read('_rels/.rels').decode('utf-8')
                    analysis['relationships'] = self.parse_relationships(rels_xml)
                
                # Analyze document structure
                if 'word/document.xml' in zipf.namelist():
                    doc_xml = zipf.read('word/document.xml').decode('utf-8')
                    analysis['document_structure'] = self.parse_document_structure(doc_xml)
                
                # Analyze numbering structure
                if 'word/numbering.xml' in zipf.namelist():
                    numbering_xml = zipf.read('word/numbering.xml').decode('utf-8')
                    analysis['numbering_structure'] = self.parse_numbering_structure(numbering_xml)
                
        except Exception as e:
            print(f"Error analyzing document {docx_path}: {e}")
            analysis['error'] = str(e)
        
        return analysis
    
    def parse_content_types(self, xml_content: str) -> Dict[str, str]:
        """Parse content types from XML"""
        try:
            root = ET.fromstring(xml_content)
            content_types = {}
            
            # Parse Default elements
            for default in root.findall('.//Default'):
                extension = default.get('Extension', '')
                content_type = default.get('ContentType', '')
                if extension and content_type:
                    content_types[f'*.{extension}'] = content_type
            
            # Parse Override elements
            for override in root.findall('.//Override'):
                part_name = override.get('PartName', '')
                content_type = override.get('ContentType', '')
                if part_name and content_type:
                    content_types[part_name] = content_type
            
            return content_types
        except Exception as e:
            return {'error': str(e)}
    
    def parse_relationships(self, xml_content: str) -> Dict[str, str]:
        """Parse relationships from XML"""
        try:
            root = ET.fromstring(xml_content)
            relationships = {}
            
            for rel in root.findall('.//Relationship'):
                rel_id = rel.get('Id', '')
                rel_type = rel.get('Type', '')
                rel_target = rel.get('Target', '')
                if rel_id and rel_type:
                    relationships[rel_id] = {'type': rel_type, 'target': rel_target}
            
            return relationships
        except Exception as e:
            return {'error': str(e)}
    
    def parse_document_structure(self, xml_content: str) -> Dict[str, Any]:
        """Parse document structure from XML"""
        try:
            root = ET.fromstring(xml_content)
            structure = {
                'namespaces': {},
                'paragraphs': [],
                'numbering_references': []
            }
            
            # Extract namespaces
            for key, value in root.attrib.items():
                if key.startswith('xmlns:'):
                    structure['namespaces'][key] = value
            
            # Count paragraphs and numbering references
            paragraphs = root.findall('.//w:p')
            structure['paragraph_count'] = len(paragraphs)
            
            for p in paragraphs:
                para_info = {'has_numbering': False, 'level': None, 'text': ''}
                
                # Check for numbering
                num_pr = p.find('.//w:numPr')
                if num_pr is not None:
                    para_info['has_numbering'] = True
                    
                    # Get numbering ID
                    num_id = num_pr.find('.//w:numId')
                    if num_id is not None:
                        para_info['num_id'] = num_id.get('w:val')
                    
                    # Get level
                    ilvl = num_pr.find('.//w:ilvl')
                    if ilvl is not None:
                        para_info['level'] = ilvl.get('w:val')
                
                # Get text content
                text_elements = p.findall('.//w:t')
                para_info['text'] = ' '.join([t.text or '' for t in text_elements])
                
                structure['paragraphs'].append(para_info)
            
            return structure
        except Exception as e:
            return {'error': str(e)}
    
    def parse_numbering_structure(self, xml_content: str) -> Dict[str, Any]:
        """Parse numbering structure from XML"""
        try:
            root = ET.fromstring(xml_content)
            structure = {
                'abstract_nums': [],
                'concrete_nums': []
            }
            
            # Parse abstract numbering definitions
            for abstract_num in root.findall('.//w:abstractNum'):
                abstract_info = {
                    'id': abstract_num.get('w:abstractNumId'),
                    'levels': []
                }
                
                for level in abstract_num.findall('.//w:lvl'):
                    level_info = {
                        'ilvl': level.get('w:ilvl'),
                        'num_fmt': None,
                        'lvl_text': None,
                        'start': None
                    }
                    
                    num_fmt = level.find('.//w:numFmt')
                    if num_fmt is not None:
                        level_info['num_fmt'] = num_fmt.get('w:val')
                    
                    lvl_text = level.find('.//w:lvlText')
                    if lvl_text is not None:
                        level_info['lvl_text'] = lvl_text.get('w:val')
                    
                    start = level.find('.//w:start')
                    if start is not None:
                        level_info['start'] = start.get('w:val')
                    
                    abstract_info['levels'].append(level_info)
                
                structure['abstract_nums'].append(abstract_info)
            
            # Parse concrete numbering instances
            for concrete_num in root.findall('.//w:num'):
                concrete_info = {
                    'id': concrete_num.get('w:numId'),
                    'abstract_num_id': None
                }
                
                abstract_num_ref = concrete_num.find('.//w:abstractNumId')
                if abstract_num_ref is not None:
                    concrete_info['abstract_num_id'] = abstract_num_ref.get('w:val')
                
                structure['concrete_nums'].append(concrete_info)
            
            return structure
        except Exception as e:
            return {'error': str(e)}
    
    def compare_documents(self, doc1_path: str, doc2_path: str) -> Dict[str, Any]:
        """Compare two Word documents and highlight differences"""
        print(f"Analyzing document 1: {doc1_path}")
        analysis1 = self.analyze_document(doc1_path)
        
        print(f"Analyzing document 2: {doc2_path}")
        analysis2 = self.analyze_document(doc2_path)
        
        comparison = {
            'file_size_difference': analysis1.get('file_size', 0) - analysis2.get('file_size', 0),
            'files_difference': set(analysis1.get('files', [])) - set(analysis2.get('files', [])),
            'content_types_difference': {},
            'structure_differences': {}
        }
        
        # Compare content types
        ct1 = analysis1.get('content_types', {})
        ct2 = analysis2.get('content_types', {})
        comparison['content_types_difference'] = {
            'only_in_doc1': set(ct1.keys()) - set(ct2.keys()),
            'only_in_doc2': set(ct2.keys()) - set(ct1.keys()),
            'different_values': {k: (ct1.get(k), ct2.get(k)) for k in set(ct1.keys()) & set(ct2.keys()) if ct1.get(k) != ct2.get(k)}
        }
        
        # Compare document structure
        doc1_structure = analysis1.get('document_structure', {})
        doc2_structure = analysis2.get('document_structure', {})
        
        comparison['structure_differences'] = {
            'paragraph_count_diff': doc1_structure.get('paragraph_count', 0) - doc2_structure.get('paragraph_count', 0),
            'namespaces_diff': set(doc1_structure.get('namespaces', {}).keys()) - set(doc2_structure.get('namespaces', {}).keys()),
            'numbered_paragraphs_doc1': sum(1 for p in doc1_structure.get('paragraphs', []) if p.get('has_numbering')),
            'numbered_paragraphs_doc2': sum(1 for p in doc2_structure.get('paragraphs', []) if p.get('has_numbering'))
        }
        
        return {
            'analysis1': analysis1,
            'analysis2': analysis2,
            'comparison': comparison
        }

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python docx_analyzer.py <docx1> <docx2>")
        print("Example: python docx_analyzer.py output/improved_accuracy_check.docx output/improved_accuracy_check-fixed2.docx")
        sys.exit(1)
    
    doc1_path = sys.argv[1]
    doc2_path = sys.argv[2]
    
    analyzer = DocxAnalyzer()
    comparison = analyzer.compare_documents(doc1_path, doc2_path)
    
    print("\n" + "="*60)
    print("DOCUMENT COMPARISON RESULTS")
    print("="*60)
    
    print(f"\nFile Size Difference: {comparison['comparison']['file_size_difference']} bytes")
    print(f"Files Only in Doc1: {comparison['comparison']['files_difference']}")
    
    print(f"\nContent Types Differences:")
    ct_diff = comparison['comparison']['content_types_difference']
    print(f"  Only in Doc1: {ct_diff['only_in_doc1']}")
    print(f"  Only in Doc2: {ct_diff['only_in_doc2']}")
    print(f"  Different Values: {ct_diff['different_values']}")
    
    print(f"\nStructure Differences:")
    struct_diff = comparison['comparison']['structure_differences']
    print(f"  Paragraph Count Diff: {struct_diff['paragraph_count_diff']}")
    print(f"  Namespaces Diff: {struct_diff['namespaces_diff']}")
    print(f"  Numbered Paragraphs - Doc1: {struct_diff['numbered_paragraphs_doc1']}")
    print(f"  Numbered Paragraphs - Doc2: {struct_diff['numbered_paragraphs_doc2']}")
    
    print(f"\nDocument 1 Analysis:")
    analysis1 = comparison['analysis1']
    print(f"  File Size: {analysis1.get('file_size', 0)} bytes")
    print(f"  Files: {len(analysis1.get('files', []))}")
    print(f"  Content Types: {len(analysis1.get('content_types', {}))}")
    
    print(f"\nDocument 2 Analysis:")
    analysis2 = comparison['analysis2']
    print(f"  File Size: {analysis2.get('file_size', 0)} bytes")
    print(f"  Files: {len(analysis2.get('files', []))}")
    print(f"  Content Types: {len(analysis2.get('content_types', {}))}")

if __name__ == "__main__":
    main() 