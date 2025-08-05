#!/usr/bin/env python3
"""
Simple Template-Based DOCX Rebuilder

This script uses a working Word document as a template and replaces
just the content while preserving the Word structure.
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

def rebuild_document_from_template(json_path: str, template_path: str, output_path: str):
    """Rebuild document using a working template"""
    try:
        print(f"Loading JSON analysis from: {json_path}")
        json_data = load_json_analysis(json_path)
        
        # Get paragraphs from JSON
        paragraphs = json_data.get('all_paragraphs', [])
        print(f"Found {len(paragraphs)} paragraphs to process")
        
        # Create new document XML
        document_xml = create_document_xml(paragraphs)
        
        # Copy template and replace document.xml
        shutil.copy2(template_path, output_path)
        
        with zipfile.ZipFile(output_path, 'a') as zipf:
            # Remove existing document.xml and add new one
            zipf.writestr('word/document.xml', document_xml)
        
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
        print("Usage: python simple_template_rebuilder.py <json_file> <template_docx> <output_docx>")
        print("Example: python simple_template_rebuilder.py output/SECTION_00_00_00_hybrid_analysis.json output/complete_accuracy_check-fixed3.docx output/simple_rebuilt.docx")
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