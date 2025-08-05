#!/usr/bin/env python3
"""
Fixed Template-Based DOCX Rebuilder

This script uses a working Word document as a template and properly
handles numbering sequences and levels.
"""

import os
import sys
import json
import zipfile
import tempfile
import shutil

def load_json_analysis(json_path: str):
    """Load the JSON analysis data"""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_document_xml(paragraphs):
    """Create the document.xml content with proper numbering"""
    body_content = ""
    current_numbering = {}  # Track numbering for each level
    
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
            
            # Update numbering for this level
            if level not in current_numbering:
                current_numbering[level] = 1
            else:
                current_numbering[level] += 1
            
            # Reset higher levels when we go to a lower level
            for higher_level in list(current_numbering.keys()):
                if higher_level > level:
                    del current_numbering[higher_level]
            
            body_content += f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:pPr>
        <w:numPr>
            <w:numId w:val="1"/>
            <w:ilvl w:val="{level}"/>
        </w:numPr>
    </w:pPr>
    <w:r>
        <w:t>{content}</w:t>
    </w:r>
</w:p>'''
        else:
            # Add regular paragraph
            body_content += f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:r>
        <w:t>{text}</w:t>
    </w:r>
</w:p>'''
    
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        {body_content}
    </w:body>
</w:document>'''

def create_numbering_xml():
    """Create proper numbering.xml with correct level definitions"""
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
            <w:numFmt w:val="decimal"/>
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
        <w:lvl w:ilvl="5">
            <w:numFmt w:val="lowerRoman"/>
            <w:lvlText w:val="%6."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="4320" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
</w:numbering>'''

def rebuild_document_from_template(json_path: str, template_path: str, output_path: str):
    """Rebuild document using a working template with proper numbering"""
    try:
        print(f"Loading JSON analysis from: {json_path}")
        json_data = load_json_analysis(json_path)
        
        # Get paragraphs from JSON
        paragraphs = json_data.get('all_paragraphs', [])
        print(f"Found {len(paragraphs)} paragraphs to process")
        
        # Create new document XML
        document_xml = create_document_xml(paragraphs)
        numbering_xml = create_numbering_xml()
        
        # Create temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Extract template to temp directory
            with zipfile.ZipFile(template_path, 'r') as zipf:
                zipf.extractall(temp_dir)
            
            # Replace document.xml
            doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            with open(doc_xml_path, 'w', encoding='utf-8') as f:
                f.write(document_xml)
            
            # Replace numbering.xml
            numbering_xml_path = os.path.join(temp_dir, 'word', 'numbering.xml')
            with open(numbering_xml_path, 'w', encoding='utf-8') as f:
                f.write(numbering_xml)
            
            # Create new ZIP file
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arc_name = os.path.relpath(file_path, temp_dir)
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
    if len(sys.argv) < 4:
        print("Usage: python fixed_template_rebuilder.py <json_file> <template_docx> <output_docx>")
        print("Example: python fixed_template_rebuilder.py output/SECTION_00_00_00_hybrid_analysis.json output/complete_accuracy_check-fixed3.docx output/fixed_rebuilt.docx")
        sys.exit(1)
    
    json_path = sys.argv[1]
    template_path = sys.argv[2]
    output_path = sys.argv[3]
    
    if not os.path.exists(json_path):
        print(f"Error: JSON file not found: {json_path}")
        sys.exit(1)
    
    if not os.path.exists(template_path):
        print(f"Error: Template file not found: {template_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    success = rebuild_document_from_template(json_path, template_path, output_path)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main() 