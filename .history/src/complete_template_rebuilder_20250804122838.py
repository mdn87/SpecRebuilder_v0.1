#!/usr/bin/env python3
"""
Complete Template-Based DOCX Rebuilder

This script creates a complete Word document with all necessary files
while preserving proper list structure and avoiding "unreadable content" warnings.
"""

import os
import sys
import json
import zipfile
import tempfile
import shutil
from datetime import datetime

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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15">
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

def create_styles_xml():
    """Create the styles.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
                <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault/>
    </w:docDefaults>
</w:styles>'''

def create_settings_xml():
    """Create the settings.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:zoom w:percent="100"/>
    <w:compat/>
</w:settings>'''

def create_web_settings_xml():
    """Create the webSettings.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:encoding w:val="x-cp1252"/>
</w:webSettings>'''

def create_font_table_xml():
    """Create the fontTable.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:font w:name="Calibri">
        <w:panose1 w:val="020F0502020204030204"/>
        <w:charset w:val="00"/>
        <w:family w:val="swiss"/>
        <w:pitch w:val="variable"/>
        <w:sig w:usb0="E00002FF" w:usb1="4000ACFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
    </w:font>
</w:fonts>'''

def create_theme_xml():
    """Create the theme1.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1>
                <a:srgbClr val="000000"/>
            </a:dk1>
            <a:lt1>
                <a:srgbClr val="FFFFFF"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="1F497D"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="EEECE1"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="4F81BD"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="C0504D"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="9BBB59"/>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="8064A2"/>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4BACC6"/>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="F79646"/>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0000FF"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="800080"/>
            </a:folHlink>
        </a:clrScheme>
    </a:themeElements>
</a:theme>'''

def create_core_properties_xml():
    """Create the core.xml content"""
    now = datetime.now().isoformat()
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>SpecRebuilder</dc:creator>
    <cp:lastModifiedBy>SpecRebuilder</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

def create_app_properties_xml():
    """Create the app.xml content"""
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>SpecRebuilder</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <LinksUpToDate>false</LinksUpToDate>
    <Pages>1</Pages>
    <Words>0</Words>
    <Characters>0</Characters>
    <Lines>0</Lines>
    <Paragraphs>0</Paragraphs>
</Properties>'''

def rebuild_document_from_json(json_path: str, output_path: str):
    """Rebuild document from JSON analysis with complete Word structure"""
    try:
        print(f"Loading JSON analysis from: {json_path}")
        json_data = load_json_analysis(json_path)
        
        # Get paragraphs from JSON
        paragraphs = json_data.get('all_paragraphs', [])
        print(f"Found {len(paragraphs)} paragraphs to process")
        
        # Create all XML content
        document_xml = create_document_xml(paragraphs)
        numbering_xml = create_numbering_xml()
        styles_xml = create_styles_xml()
        settings_xml = create_settings_xml()
        web_settings_xml = create_web_settings_xml()
        font_table_xml = create_font_table_xml()
        theme_xml = create_theme_xml()
        core_properties_xml = create_core_properties_xml()
        app_properties_xml = create_app_properties_xml()
        
        # Create temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create document structure
            doc_dir = os.path.join(temp_dir, 'word')
            theme_dir = os.path.join(doc_dir, 'theme')
            doc_props_dir = os.path.join(temp_dir, 'docProps')
            
            os.makedirs(doc_dir, exist_ok=True)
            os.makedirs(theme_dir, exist_ok=True)
            os.makedirs(doc_props_dir, exist_ok=True)
            
            # Create all XML files
            xml_files = [
                (os.path.join(doc_dir, 'document.xml'), document_xml),
                (os.path.join(doc_dir, 'numbering.xml'), numbering_xml),
                (os.path.join(doc_dir, 'styles.xml'), styles_xml),
                (os.path.join(doc_dir, 'settings.xml'), settings_xml),
                (os.path.join(doc_dir, 'webSettings.xml'), web_settings_xml),
                (os.path.join(doc_dir, 'fontTable.xml'), font_table_xml),
                (os.path.join(theme_dir, 'theme1.xml'), theme_xml),
                (os.path.join(doc_props_dir, 'core.xml'), core_properties_xml),
                (os.path.join(doc_props_dir, 'app.xml'), app_properties_xml)
            ]
            
            for file_path, content in xml_files:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
            
            # Create [Content_Types].xml
            content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
            
            with open(os.path.join(temp_dir, '[Content_Types].xml'), 'w', encoding='utf-8') as f:
                f.write(content_types)
            
            # Create _rels/.rels
            rels_dir = os.path.join(temp_dir, '_rels')
            os.makedirs(rels_dir, exist_ok=True)
            
            rels_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
            
            with open(os.path.join(rels_dir, '.rels'), 'w', encoding='utf-8') as f:
                f.write(rels_content)
            
            # Create word/_rels/document.xml.rels
            word_rels_dir = os.path.join(doc_dir, '_rels')
            os.makedirs(word_rels_dir, exist_ok=True)
            
            word_rels_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
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
                    ('word/numbering.xml', os.path.join(temp_dir, 'word', 'numbering.xml')),
                    ('word/styles.xml', os.path.join(temp_dir, 'word', 'styles.xml')),
                    ('word/settings.xml', os.path.join(temp_dir, 'word', 'settings.xml')),
                    ('word/webSettings.xml', os.path.join(temp_dir, 'word', 'webSettings.xml')),
                    ('word/fontTable.xml', os.path.join(temp_dir, 'word', 'fontTable.xml')),
                    ('word/theme/theme1.xml', os.path.join(temp_dir, 'word', 'theme', 'theme1.xml')),
                    ('docProps/core.xml', os.path.join(temp_dir, 'docProps', 'core.xml')),
                    ('docProps/app.xml', os.path.join(temp_dir, 'docProps', 'app.xml'))
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
        print("Usage: python complete_template_rebuilder.py <json_file> <output_docx>")
        print("Example: python complete_template_rebuilder.py output/SECTION_00_00_00_hybrid_analysis.json output/complete_template_rebuilt.docx")
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