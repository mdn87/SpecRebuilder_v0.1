#!/usr/bin/env python3
"""
Template-Based DOCX Rebuilder

This script uses a working Word document as a template and injects
our content while preserving the proper Word structure.
"""

import os
import sys
import json
import zipfile
import tempfile
import shutil
from docx import Document
from docx.oxml import parse_xml

def load_json_analysis(json_path: str):
    """Load the JSON analysis data"""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_numbered_paragraph_xml(text, level=0):
    """Create XML for a numbered paragraph"""
    return f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:pPr>
        <w:numPr>
            <w:numId w:val="1"/>
            <w:ilvl w:val="{level}"/>
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>{text}</w:t>
    </w:r>
</w:p>'''

def create_regular_paragraph_xml(text):
    """Create XML for a regular paragraph"""
    return f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:r>
        <w:t>{text}</w:t>
    </w:r>
</w:p>'''

def create_document_xml(paragraphs):
    """Create the document.xml content"""
    body_content = ""
    
    for para_data in paragraphs:
        text = para_data.get('text', '').strip()
        if not text:
            continue
        
        # Check if this paragraph has numbering
        has_numbering = bool(para_data.get('list_number') or para_data.get('inferred_number'))
        
        if has_numbering:
            # Use cleaned content if available, otherwise use original text
            content = para_data.get('cleaned_content', text)
            level = para_data.get('level', 0)
            body_content += create_numbered_paragraph_xml(content, level)
        else:
            # Add regular paragraph
            body_content += create_regular_paragraph_xml(text)
    
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        {body_content}
    </w:body>
</w:document>'''

def create_numbering_xml():
    """Create the numbering.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="1">
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1.%2."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="1440" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="2">
            <w:numFmt w:val="lowerLetter"/>
            <w:lvlText w:val="%3."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="2160" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="3">
            <w:numFmt w:val="lowerLetter"/>
            <w:lvlText w:val="%4."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="2880" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
        <w:lvl w:ilvl="4">
            <w:numFmt w:val="lowerLetter"/>
            <w:lvlText w:val="%5."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="3600" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
</w:numbering>'''

def rebuild_document_from_json(json_path: str, output_path: str, template_path: str = None):
    """Rebuild document from JSON analysis using template approach"""
    try:
        print(f"Loading JSON analysis from: {json_path}")
        json_data = load_json_analysis(json_path)
        
        # Get paragraphs from JSON
        paragraphs = json_data.get('all_paragraphs', [])
        print(f"Found {len(paragraphs)} paragraphs to process")
        
        # Create document XML
        document_xml = create_document_xml(paragraphs)
        numbering_xml = create_numbering_xml()
        
        # Create temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create document structure
            doc_dir = os.path.join(temp_dir, 'word')
            os.makedirs(doc_dir, exist_ok=True)
            
            # Create document.xml
            with open(os.path.join(doc_dir, 'document.xml'), 'w', encoding='utf-8') as f:
                f.write(document_xml)
            
            # Create numbering.xml
            with open(os.path.join(doc_dir, 'numbering.xml'), 'w', encoding='utf-8') as f:
                f.write(numbering_xml)
            
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
            
            # Create _rels/.rels
            rels_dir = os.path.join(temp_dir, '_rels')
            os.makedirs(rels_dir, exist_ok=True)
            
            rels_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
            
            with open(os.path.join(rels_dir, '.rels'), 'w', encoding='utf-8') as f:
                f.write(rels_content)
            
            # Create word/_rels/document.xml.rels
            word_rels_dir = os.path.join(doc_dir, '_rels')
            os.makedirs(word_rels_dir, exist_ok=True)
            
            word_rels_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''
            
            with open(os.path.join(word_rels_dir, 'document.xml.rels'), 'w', encoding='utf-8') as f:
                f.write(word_rels_content)
            
            # Create ZIP file
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                files_to_add = [
                    ('[Content_Types].xml', os.path.join(temp_dir, '[Content_Types].xml')),
                    ('_rels/.rels', os.path.join(temp_dir, '_rels', '.rels')),
                    ('word/document.xml', os.path.join(temp_dir, 'word', 'document.xml')),
                    ('word/_rels/document.xml.rels', os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')),
                    ('word/numbering.xml', os.path.join(temp_dir, 'word', 'numbering.xml'))
                ]
                
                for arc_name, file_path in files_to_add:
                    if os.path.exists(file_path):
                        zipf.write(file_path, arc_name)
        
        print(f"Document saved to: {output_path}")
        print("Document rebuild complete!")
        return True
        
    except Exception as e:
        print(f"Error rebuilding document: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python template_based_rebuilder.py <json_file> <output_docx>")
        print("Example: python template_based_rebuilder.py output/SECTION_00_00_00_hybrid_analysis.json output/template_rebuilt.docx")
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
    
    success = rebuild_document_from_json(json_path, output_path)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main() 