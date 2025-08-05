#!/usr/bin/env python3
"""
Enhanced List Reconstructor

This script creates Word documents with proper multilevel lists
using COM API with custom list templates.
"""

import os
import sys
import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

# Try to import win32com
try:
    import win32com.client
    import pythoncom
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("Warning: win32com not available. Install with: pip install pywin32")

@dataclass
class ParagraphData:
    """Represents a paragraph with numbering information"""
    index: int
    list_number: str
    text: str
    level: Optional[int] = None
    inferred_number: Optional[str] = None
    cleaned_content: Optional[str] = None

class EnhancedListReconstructor:
    """Enhanced Word document reconstructor with proper multilevel lists"""
    
    def __init__(self):
        if not WIN32COM_AVAILABLE:
            raise ImportError("win32com not available. Install with: pip install pywin32")
        
        # Word constants
        self.WD_LIST_NUMBER_STYLE_ARABIC = 1
        self.WD_LIST_NUMBER_STYLE_UPPER_LETTER = 2
        self.WD_LIST_NUMBER_STYLE_LOWER_LETTER = 3
        self.WD_LIST_NUMBER_STYLE_UPPER_ROMAN = 4
        self.WD_LIST_NUMBER_STYLE_LOWER_ROMAN = 5
        self.WD_LIST_NUMBER_STYLE_DECIMAL = 6
        
        # COM constants
        self.WD_LIST_APPLY_TO_WHOLE_LIST = 1
        self.WD_LIST_CONTINUE_PREVIOUS_LIST = True
        self.WD_LIST_NEW_LIST = False
    
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
    
    def determine_numbering_style(self, numbering: str) -> int:
        """Determine the Word numbering style constant"""
        if not numbering:
            return self.WD_LIST_NUMBER_STYLE_ARABIC
        
        # Decimal patterns: 1.0, 1.01, 2.0, etc.
        if re.match(r'^\d+\.\d+$', numbering):
            return self.WD_LIST_NUMBER_STYLE_DECIMAL
        
        # Simple decimal: 1., 2., 3., etc.
        if re.match(r'^\d+\.$', numbering):
            return self.WD_LIST_NUMBER_STYLE_ARABIC
        
        # Upper case letters: A., B., C., etc.
        if re.match(r'^[A-Z]\.$', numbering):
            return self.WD_LIST_NUMBER_STYLE_UPPER_LETTER
        
        # Lower case letters: a., b., c., etc.
        if re.match(r'^[a-z]\.$', numbering):
            return self.WD_LIST_NUMBER_STYLE_LOWER_LETTER
        
        # Roman numerals (basic detection)
        if re.match(r'^[IVX]+\.$', numbering):
            return self.WD_LIST_NUMBER_STYLE_UPPER_ROMAN
        
        # Default to decimal
        return self.WD_LIST_NUMBER_STYLE_ARABIC
    
    def create_custom_list_template(self, word_app, levels_config: List[Dict]):
        """Create a custom list template with specific level configurations"""
        try:
            # Get the numbering gallery
            gallery = word_app.ListGalleries(1)  # wdListGalleryNumbering
            
            # Create a new list template
            list_template = gallery.ListTemplates(1)  # Use first template as base
            
            # Configure each level
            for level_idx, config in enumerate(levels_config, 1):
                if level_idx <= list_template.ListLevels.Count:
                    level = list_template.ListLevels(level_idx)
                    level.NumberingStyle = config.get('style', self.WD_LIST_NUMBER_STYLE_ARABIC)
                    
                    # Set level text format
                    if config.get('format'):
                        level.NumberFormat = config['format']
                    
                    # Set alignment
                    if config.get('alignment'):
                        level.Alignment = config['alignment']
            
            return list_template
            
        except Exception as e:
            print(f"Warning: Could not create custom list template: {e}")
            # Fall back to default template
            return word_app.ListGalleries(1).ListTemplates(1)
    
    def create_word_document_with_lists(self, paragraphs: List[ParagraphData], output_path: str):
        """Create a new Word document with proper multilevel lists"""
        # Initialize COM
        pythoncom.CoInitialize()
        word = None
        doc = None
        
        try:
            # Create Word application
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # Create new document
            doc = word.Documents.Add()
            
            # Analyze numbering patterns to create appropriate list template
            levels_config = self.analyze_numbering_patterns(paragraphs)
            list_template = self.create_custom_list_template(word, levels_config)
            
            # Track current list state
            current_list = None
            list_objects = {}  # Store list objects for each level
            
            for para_data in paragraphs:
                # Skip empty paragraphs
                if not para_data.text.strip():
                    continue
                
                # Add paragraph
                para = doc.Paragraphs.Add()
                
                # Determine if this should be numbered
                has_numbering = bool(para_data.list_number or para_data.inferred_number)
                
                if has_numbering:
                    # Get the numbering to use
                    numbering = para_data.list_number if para_data.list_number else para_data.inferred_number
                    level = para_data.level if para_data.level is not None else 0
                    
                    # Set the text content
                    if para_data.cleaned_content:
                        para.Range.Text = para_data.cleaned_content
                    else:
                        para.Range.Text = para_data.text
                    
                    # Apply numbering with proper list formatting
                    self.apply_list_numbering(para, numbering, level, list_template, list_objects)
                    
                else:
                    # Regular paragraph (no numbering)
                    para.Range.Text = para_data.text
                
                # Add line break
                para.Range.InsertAfter("\n")
            
            # Save the document
            doc.SaveAs(os.path.abspath(output_path))
            print(f"Document saved to: {output_path}")
            
        finally:
            # Clean up
            if doc is not None:
                try:
                    doc.Close(True)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    
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
                        'alignment': 0  # Left alignment
                    }
        
        # Convert to list configuration
        levels_config = []
        for level in sorted(level_patterns.keys()):
            levels_config.append(level_patterns[level])
        
        return levels_config
    
    def apply_list_numbering(self, para, numbering: str, level: int, list_template, list_objects: Dict):
        """Apply proper list numbering to a paragraph"""
        try:
            # Determine if we need a new list or continue existing
            list_key = f"level_{level}"
            
            if list_key not in list_objects:
                # Start new list for this level
                list_obj = para.Range.ListFormat
                list_obj.ApplyListTemplate(
                    ListTemplate=list_template,
                    ContinuePreviousList=False,
                    ApplyTo=self.WD_LIST_APPLY_TO_WHOLE_LIST
                )
                list_objects[list_key] = list_obj
            else:
                # Continue existing list
                para.Range.ListFormat = list_objects[list_key]
            
            # Set the specific list level
            para.Range.ListFormat.ListLevelNumber = level + 1
            
            # Apply the numbering style for this level
            numbering_style = self.determine_numbering_style(numbering)
            para.Range.ListFormat.ListLevels(level + 1).NumberingStyle = numbering_style
            
        except Exception as e:
            print(f"Warning: Could not apply list numbering '{numbering}' at level {level}: {e}")
            # Fall back to plain text with indentation
            indent = "  " * level
            para.Range.Text = f"{indent}{numbering} {para.Range.Text}"
    
    def reconstruct_document(self, json_path: str, output_path: str):
        """Main method to reconstruct a Word document from JSON analysis"""
        print(f"Loading JSON analysis from: {json_path}")
        json_data = self.load_json_analysis(json_path)
        
        print(f"Parsing {len(json_data.get('all_paragraphs', []))} paragraphs...")
        paragraphs = self.parse_paragraphs_from_json(json_data)
        
        print(f"Creating Word document with {len(paragraphs)} paragraphs...")
        self.create_word_document_with_lists(paragraphs, output_path)
        
        print("Document reconstruction complete!")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python enhanced_list_reconstructor.py <json_file> <output_docx>")
        print("Example: python enhanced_list_reconstructor.py output/SECTION_00_00_00_hybrid_analysis.json enhanced_reconstructed_SECTION_00_00_00.docx")
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
        reconstructor = EnhancedListReconstructor()
        reconstructor.reconstruct_document(json_path, output_path)
        
    except ImportError as e:
        print(f"Error: {e}")
        print("Please install pywin32: pip install pywin32")
        sys.exit(1)
    except Exception as e:
        print(f"Error in document reconstruction: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 