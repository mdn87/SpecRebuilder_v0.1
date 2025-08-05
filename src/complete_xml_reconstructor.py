#!/usr/bin/env python3
"""
Complete XML List Reconstructor

This script creates Word documents with proper multilevel lists
by directly manipulating the Word document's XML structure.
Includes all necessary Word document components for full compatibility.
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
from datetime import datetime

@dataclass
class ParagraphData:
    """Represents a paragraph with numbering information"""
    index: int
    list_number: str
    text: str
    level: Optional[int] = None
    inferred_number: Optional[str] = None
    cleaned_content: Optional[str] = None

class CompleteXMLReconstructor:
    """Complete XML-based Word document reconstructor with all necessary components"""
    
    def __init__(self):
        # Word XML namespaces - using standard Word 2016+ namespaces
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/',
            'dcmitype': 'http://purl.org/dc/dcmitype/',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
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
        """Create the numbering.xml content with proper Word structure"""
        # Create the numbering XML structure
        numbering = ET.Element('w:numbering')
        
        # Add namespace
        numbering.set('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        
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
            
            # Add start value for proper numbering
            start = ET.SubElement(level, 'w:start')
            start.set('w:val', '1')
        
        # Create concrete numbering instance
        num = ET.SubElement(numbering, 'w:num')
        num.set('w:numId', '1')
        
        abstract_num_ref = ET.SubElement(num, 'w:abstractNumId')
        abstract_num_ref.set('w:val', '0')
        
        return ET.tostring(numbering, encoding='unicode', xml_declaration=True)
    
    def create_document_xml(self, paragraphs: List[ParagraphData]) -> str:
        """Create the document.xml content with proper list formatting"""
        # Create the document XML structure
        document = ET.Element('w:document')
        
        # Add all namespaces
        for prefix, uri in self.namespaces.items():
            document.set(f'xmlns:{prefix}', uri)
        
        # Add compatibility settings
        mc = ET.SubElement(document, 'mc:Ignorable')
        mc.set('w14:val', 'http://schemas.microsoft.com/office/word/2010/wordml')
        mc.set('w15:val', 'http://schemas.microsoft.com/office/word/2012/wordml')
        
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
            
            # Add text run with proper properties
            r = ET.SubElement(p, 'w:r')
            
            # Add run properties
            r_pr = ET.SubElement(r, 'w:rPr')
            
            # Add text
            t = ET.SubElement(r, 'w:t')
            t.set('xml:space', 'preserve')  # Preserve whitespace
            
            # Set the text content
            if para_data.cleaned_content:
                t.text = para_data.cleaned_content
            else:
                t.text = para_data.text
        
        return ET.tostring(document, encoding='unicode', xml_declaration=True)
    
    def create_styles_xml(self) -> str:
        """Create the styles.xml content"""
        styles = ET.Element('w:styles')
        styles.set('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        
        # Add default paragraph style
        doc_defaults = ET.SubElement(styles, 'w:docDefaults')
        p_pr_default = ET.SubElement(doc_defaults, 'w:pPrDefault')
        p_pr = ET.SubElement(p_pr_default, 'w:pPr')
        
        # Add default run style
        r_pr_default = ET.SubElement(doc_defaults, 'w:rPrDefault')
        r_pr = ET.SubElement(r_pr_default, 'w:rPr')
        
        return ET.tostring(styles, encoding='unicode', xml_declaration=True)
    
    def create_settings_xml(self) -> str:
        """Create the settings.xml content"""
        settings = ET.Element('w:settings')
        settings.set('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        
        # Add compatibility settings
        compat = ET.SubElement(settings, 'w:compat')
        
        # Add zoom settings
        zoom = ET.SubElement(settings, 'w:zoom')
        zoom.set('w:percent', '100')
        
        return ET.tostring(settings, encoding='unicode', xml_declaration=True)
    
    def create_web_settings_xml(self) -> str:
        """Create the webSettings.xml content"""
        web_settings = ET.Element('w:webSettings')
        web_settings.set('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        
        # Add encoding
        encoding = ET.SubElement(web_settings, 'w:encoding')
        encoding.set('w:val', 'utf-8')
        
        return ET.tostring(web_settings, encoding='unicode', xml_declaration=True)
    
    def create_font_table_xml(self) -> str:
        """Create the fontTable.xml content"""
        font_table = ET.Element('w:fonts')
        font_table.set('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        
        # Add default font
        font = ET.SubElement(font_table, 'w:font')
        font.set('w:name', 'Calibri')
        
        # Add font family
        family = ET.SubElement(font, 'w:family')
        family.set('w:val', 'swiss')
        
        # Add pitch
        pitch = ET.SubElement(font, 'w:pitch')
        pitch.set('w:val', 'variable')
        
        return ET.tostring(font_table, encoding='unicode', xml_declaration=True)
    
    def create_theme_xml(self) -> str:
        """Create the theme1.xml content"""
        theme = ET.Element('a:theme')
        theme.set('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        theme.set('name', 'Office Theme')
        
        # Add theme elements
        theme_elements = ET.SubElement(theme, 'a:themeElements')
        
        # Add color scheme
        clr_scheme = ET.SubElement(theme_elements, 'a:clrScheme')
        clr_scheme.set('name', 'Office')
        
        return ET.tostring(theme, encoding='unicode', xml_declaration=True)
    
    def create_core_properties_xml(self) -> str:
        """Create the core.xml content"""
        core_properties = ET.Element('cp:coreProperties')
        
        # Add namespaces
        for prefix, uri in self.namespaces.items():
            if prefix in ['cp', 'dc', 'dcterms', 'dcmitype', 'xsi']:
                core_properties.set(f'xmlns:{prefix}', uri)
        
        # Add creation date
        created = ET.SubElement(core_properties, 'dcterms:created')
        created.set('xsi:type', 'dcterms:W3CDTF')
        created.text = datetime.now().isoformat()
        
        # Add modification date
        modified = ET.SubElement(core_properties, 'dcterms:modified')
        modified.set('xsi:type', 'dcterms:W3CDTF')
        modified.text = datetime.now().isoformat()
        
        return ET.tostring(core_properties, encoding='unicode', xml_declaration=True)
    
    def create_app_properties_xml(self) -> str:
        """Create the app.xml content"""
        properties = ET.Element('Properties')
        properties.set('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
        properties.set('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
        
        # Add application
        application = ET.SubElement(properties, 'Application')
        application.text = 'Microsoft Office Word'
        
        # Add document security
        doc_security = ET.SubElement(properties, 'DocSecurity')
        doc_security.text = '0'
        
        # Add scale crop
        scale_crop = ET.SubElement(properties, 'ScaleCrop')
        scale_crop.text = 'false'
        
        # Add links up to date
        links_up_to_date = ET.SubElement(properties, 'LinksUpToDate')
        links_up_to_date.text = 'false'
        
        # Add pages
        pages = ET.SubElement(properties, 'Pages')
        pages.text = '1'
        
        # Add words
        words = ET.SubElement(properties, 'Words')
        words.text = '0'
        
        # Add characters
        characters = ET.SubElement(properties, 'Characters')
        characters.text = '0'
        
        # Add lines
        lines = ET.SubElement(properties, 'Lines')
        lines.text = '0'
        
        # Add paragraphs
        paragraphs = ET.SubElement(properties, 'Paragraphs')
        paragraphs.text = '0'
        
        return ET.tostring(properties, encoding='unicode', xml_declaration=True)
    
    def create_word_document_xml(self, paragraphs: List[ParagraphData], output_path: str):
        """Create a new Word document using complete XML structure"""
        # Analyze numbering patterns
        levels_config = self.analyze_numbering_patterns(paragraphs)
        
        # Create temporary directory for document structure
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create document structure
            doc_dir = os.path.join(temp_dir, 'word')
            os.makedirs(doc_dir, exist_ok=True)
            
            # Create theme directory
            theme_dir = os.path.join(doc_dir, 'theme')
            os.makedirs(theme_dir, exist_ok=True)
            
            # Create docProps directory
            doc_props_dir = os.path.join(temp_dir, 'docProps')
            os.makedirs(doc_props_dir, exist_ok=True)
            
            # Create numbering.xml
            numbering_xml = self.create_numbering_xml(levels_config)
            with open(os.path.join(doc_dir, 'numbering.xml'), 'w', encoding='utf-8') as f:
                f.write(numbering_xml)
            
            # Create document.xml
            document_xml = self.create_document_xml(paragraphs)
            with open(os.path.join(doc_dir, 'document.xml'), 'w', encoding='utf-8') as f:
                f.write(document_xml)
            
            # Create styles.xml
            styles_xml = self.create_styles_xml()
            with open(os.path.join(doc_dir, 'styles.xml'), 'w', encoding='utf-8') as f:
                f.write(styles_xml)
            
            # Create settings.xml
            settings_xml = self.create_settings_xml()
            with open(os.path.join(doc_dir, 'settings.xml'), 'w', encoding='utf-8') as f:
                f.write(settings_xml)
            
            # Create webSettings.xml
            web_settings_xml = self.create_web_settings_xml()
            with open(os.path.join(doc_dir, 'webSettings.xml'), 'w', encoding='utf-8') as f:
                f.write(web_settings_xml)
            
            # Create fontTable.xml
            font_table_xml = self.create_font_table_xml()
            with open(os.path.join(doc_dir, 'fontTable.xml'), 'w', encoding='utf-8') as f:
                f.write(font_table_xml)
            
            # Create theme1.xml
            theme_xml = self.create_theme_xml()
            with open(os.path.join(theme_dir, 'theme1.xml'), 'w', encoding='utf-8') as f:
                f.write(theme_xml)
            
            # Create core.xml
            core_xml = self.create_core_properties_xml()
            with open(os.path.join(doc_props_dir, 'core.xml'), 'w', encoding='utf-8') as f:
                f.write(core_xml)
            
            # Create app.xml
            app_xml = self.create_app_properties_xml()
            with open(os.path.join(doc_props_dir, 'app.xml'), 'w', encoding='utf-8') as f:
                f.write(app_xml)
            
            # Create [Content_Types].xml with complete content types
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
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>'''
            
            with open(os.path.join(temp_dir, '[Content_Types].xml'), 'w', encoding='utf-8') as f:
                f.write(content_types)
            
            # Create _rels directory and .rels file
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
            
            # Create word/_rels directory and document.xml.rels
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
            
            # Create ZIP file (Word document) with proper compression
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zipf:
                # Add all files to the ZIP in proper order
                files_to_add = [
                    ('[Content_Types].xml', os.path.join(temp_dir, '[Content_Types].xml')),
                    ('_rels/.rels', os.path.join(temp_dir, '_rels', '.rels')),
                    ('docProps/core.xml', os.path.join(temp_dir, 'docProps', 'core.xml')),
                    ('docProps/app.xml', os.path.join(temp_dir, 'docProps', 'app.xml')),
                    ('word/document.xml', os.path.join(temp_dir, 'word', 'document.xml')),
                    ('word/_rels/document.xml.rels', os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')),
                    ('word/numbering.xml', os.path.join(temp_dir, 'word', 'numbering.xml')),
                    ('word/styles.xml', os.path.join(temp_dir, 'word', 'styles.xml')),
                    ('word/settings.xml', os.path.join(temp_dir, 'word', 'settings.xml')),
                    ('word/webSettings.xml', os.path.join(temp_dir, 'word', 'webSettings.xml')),
                    ('word/fontTable.xml', os.path.join(temp_dir, 'word', 'fontTable.xml')),
                    ('word/theme/theme1.xml', os.path.join(temp_dir, 'word', 'theme', 'theme1.xml'))
                ]
                
                for arc_name, file_path in files_to_add:
                    if os.path.exists(file_path):
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
        
        print(f"Creating complete Word document with {len(paragraphs)} paragraphs...")
        self.create_word_document_xml(paragraphs, output_path)
        
        print("Document reconstruction complete!")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python complete_xml_reconstructor.py <json_file> <output_docx>")
        print("Example: python complete_xml_reconstructor.py output/SECTION_00_00_00_hybrid_analysis.json complete_reconstructed_SECTION_00_00_00.docx")
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
        reconstructor = CompleteXMLReconstructor()
        reconstructor.reconstruct_document(json_path, output_path)
        
    except Exception as e:
        print(f"Error in document reconstruction: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 