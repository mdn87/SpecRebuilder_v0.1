#!/usr/bin/env python3
"""
XML List Reconstructor

This script creates Word documents with proper multilevel lists
by directly manipulating the Word document's XML structure.
"""

import os
import sys
import json
import re
import zipfile
import tempfile
import shutil
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from xml.etree import ElementTree as ET

@dataclass
class ParagraphData:
    """Represents a paragraph with numbering information"""
    index: int
    list_number: str
    text: str
    level: Optional[int] = None
    inferred_number: Optional[str] = None
    cleaned_content: Optional[str] = None

class XMLListReconstructor:
    """XML-based Word document reconstructor with proper multilevel lists"""
    
    def __init__(self):
        # Word XML namespaces
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        }
        
        # Numbering style mappings
        self.numbering_styles = {
            'decimal': 'decimal',
            'upperLetter': 'upperLetter',
            'lowerLetter': 'lowerLetter',
            'upperRoman': 'upperRoman',
            'lowerRoman': 'lowerRoman',
            'bullet': 'bullet'
        }
    
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
    
    def determine_numbering_style(self, numbering: str) -> str:
        """Determine the numbering style based on the numbering pattern"""
        if not numbering:
            return 'decimal'
        
        # Decimal patterns: 1.0, 1.01, 2.0, etc.
        if re.match(r'^\d+\.\d+$', numbering):
            return 'decimal'
        
        # Simple decimal: 1., 2., 3., etc.
        if re.match(r'^\d+\.$', numbering):
            return 'decimal'
        
        # Upper case letters: A., B., C., etc.
        if re.match(r'^[A-Z]\.$', numbering):
            return 'upperLetter'
        
        # Lower case letters: a., b., c., etc.
        if re.match(r'^[a-z]\.$', numbering):
            return 'lowerLetter'
        
        # Roman numerals (basic detection)
        if re.match(r'^[IVX]+\.$', numbering):
            return 'upperRoman'
        
        # Default to decimal
        return 'decimal'
    
    def create_numbering_xml(self, levels_config: List[Dict]) -> str:
        """Create the numbering.xml content with custom list definitions"""
        # Create the numbering XML structure
        numbering = ET.Element('w:numbering')
        
        # Create abstract numbering definition
        abstract_num = ET.SubElement(numbering, 'w:abstractNum')
        abstract_num.set('w:abstractNumId', '0')
        
        # Add level definitions
        for level_idx, config in enumerate(levels_config):
            level = ET.SubElement(abstract_num, 'w:lvl')
            level.set('w:ilvl', str(level_idx))
            
            # Set numbering format
            num_fmt = ET.SubElement(level, 'w:numFmt')
            num_fmt.set('w:val', config.get('style', 'decimal'))
            
            # Set level text
            lvl_text = ET.SubElement(level, 'w:lvlText')
            lvl_text.set('w:val', config.get('format', f'%{level_idx + 1}.'))
            
            # Set level justification
            lvl_jc = ET.SubElement(level, 'w:lvlJc')
            lvl_jc.set('w:val', 'left')
            
            # Set level indentation
            p_pr = ET.SubElement(level, 'w:pPr')
            ind = ET.SubElement(p_pr, 'w:ind')
            ind.set('w:left', str(level_idx * 720))  # 720 twips = 0.5 inch
            ind.set('w:hanging', '360')  # 360 twips = 0.25 inch
        
        # Create concrete numbering instance
        num = ET.SubElement(numbering, 'w:num')
        num.set('w:numId', '1')
        
        abstract_num_ref = ET.SubElement(num, 'w:abstractNumId')
        abstract_num_ref.set('w:val', '0')
        
        return ET.tostring(numbering, encoding='unicode')
    
    def create_document_xml(self, paragraphs: List[ParagraphData]) -> str:
        """Create the document.xml content with proper list formatting"""
        # Create the document XML structure
        document = ET.Element('w:document')
        
        # Add namespaces
        for prefix, uri in self.namespaces.items():
            document.set(f'xmlns:{prefix}', uri)
        
        # Create body
        body = ET.SubElement(document, 'w:body')
        
        # Add paragraphs
        for para_data in paragraphs:
            if not para_data.text.strip():
                continue
            
            # Create paragraph
            p = ET.SubElement(body, 'w:p')
            
            # Add paragraph properties
            p_pr = ET.SubElement(p, 'w:pPr')
            
            # Determine if this should be numbered
            has_numbering = bool(para_data.list_number or para_data.inferred_number)
            
            if has_numbering:
                # Add numbering properties
                num_pr = ET.SubElement(p_pr, 'w:numPr')
                
                # Set numbering ID
                num_id = ET.SubElement(num_pr, 'w:numId')
                num_id.set('w:val', '1')
                
                # Set level
                ilvl = ET.SubElement(num_pr, 'w:ilvl')
                level = para_data.level if para_data.level is not None else 0
                ilvl.set('w:val', str(level))
            
            # Add text run
            r = ET.SubElement(p, 'w:r')
            t = ET.SubElement(r, 'w:t')
            
            # Set the text content
            if para_data.cleaned_content:
                t.text = para_data.cleaned_content
            else:
                t.text = para_data.text
        
        return ET.tostring(document, encoding='unicode')
    
    def create_word_document_xml(self, paragraphs: List[ParagraphData], output_path: str):
        """Create a new Word document using XML structure"""
        # Analyze numbering patterns
        levels_config = self.analyze_numbering_patterns(paragraphs)
        
        # Create temporary directory for document structure
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create document structure
            doc_dir = os.path.join(temp_dir, 'word')
            os.makedirs(doc_dir, exist_ok=True)
            
            # Create numbering.xml
            numbering_xml = self.create_numbering_xml(levels_config)
            with open(os.path.join(doc_dir, 'numbering.xml'), 'w', encoding='utf-8') as f:
                f.write(numbering_xml)
            
            # Create document.xml
            document_xml = self.create_document_xml(paragraphs)
            with open(os.path.join(doc_dir, 'document.xml'), 'w', encoding='utf-8') as f:
                f.write(document_xml)
            
            # Create [Content_Types].xml
            content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''
            
            with open(os.path.join(temp_dir, '[Content_Types].xml'), 'w', encoding='utf-8') as f:
                f.write(content_types)
            
            # Create _rels directory and .rels file
            rels_dir = os.path.join(temp_dir, '_rels')
            os.makedirs(rels_dir, exist_ok=True)
            
            rels_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
            
            with open(os.path.join(rels_dir, '.rels'), 'w', encoding='utf-8') as f:
                f.write(rels_content)
            
            # Create word/_rels directory and document.xml.rels
            word_rels_dir = os.path.join(doc_dir, '_rels')
            os.makedirs(word_rels_dir, exist_ok=True)
            
            word_rels_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''
            
            with open(os.path.join(word_rels_dir, 'document.xml.rels'), 'w', encoding='utf-8') as f:
                f.write(word_rels_content)
            
            # Create ZIP file (Word document)
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Add all files to the ZIP
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arc_name)
        
        print(f"Document saved to: {output_path}")
    
    def analyze_numbering_patterns(self, paragraphs: List[ParagraphData]) -> List[Dict]:
        """Analyze numbering patterns to determine list level configurations"""
        level_patterns = {}
        
        for para in paragraphs:
            if para.level is not None:
                numbering = para.list_number or para.inferred_number
                if numbering:
                    style = self.determine_numbering_style(numbering)
                    level_patterns[para.level] = {
                        'style': style,
                        'format': numbering,
                        'alignment': 'left'
                    }
        
        # Convert to list configuration
        levels_config = []
        for level in sorted(level_patterns.keys()):
            levels_config.append(level_patterns[level])
        
        return levels_config
    
    def reconstruct_document(self, json_path: str, output_path: str):
        """Main method to reconstruct a Word document from JSON analysis"""
        print(f"Loading JSON analysis from: {json_path}")
        json_data = self.load_json_analysis(json_path)
        
        print(f"Parsing {len(json_data.get('all_paragraphs', []))} paragraphs...")
        paragraphs = self.parse_paragraphs_from_json(json_data)
        
        print(f"Creating Word document with {len(paragraphs)} paragraphs...")
        self.create_word_document_xml(paragraphs, output_path)
        
        print("Document reconstruction complete!")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python xml_list_reconstructor.py <json_file> <output_docx>")
        print("Example: python xml_list_reconstructor.py output/SECTION_00_00_00_hybrid_analysis.json xml_reconstructed_SECTION_00_00_00.docx")
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
        reconstructor = XMLListReconstructor()
        reconstructor.reconstruct_document(json_path, output_path)
        
    except Exception as e:
        print(f"Error in document reconstruction: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 