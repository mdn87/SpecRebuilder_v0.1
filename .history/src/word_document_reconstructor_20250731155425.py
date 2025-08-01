#!/usr/bin/env python3
"""
Word Document Reconstructor

This script takes JSON analysis data from the enhanced hybrid detector
and reconstructs a Word document with proper list levels.
"""

import os
import sys
import json
import re
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

# Try to import win32com, but provide fallback if not available
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

class WordDocumentReconstructor:
    """Reconstructs Word documents from JSON analysis data"""
    
    def __init__(self):
        if not WIN32COM_AVAILABLE:
            raise ImportError("win32com not available. Install with: pip install pywin32")
        
        # Word constants
        self.WD_LIST_NUMBER_STYLE_ARABIC = 1
        self.WD_LIST_NUMBER_STYLE_UPPER_LETTER = 2
        self.WD_LIST_NUMBER_STYLE_LOWER_LETTER = 3
        self.WD_LIST_NUMBER_STYLE_UPPER_ROMAN = 4
        self.WD_LIST_NUMBER_STYLE_LOWER_ROMAN = 5
        
        # Numbering style mapping
        self.numbering_style_map = {
            'decimal': self.WD_LIST_NUMBER_STYLE_ARABIC,
            'upper_letter': self.WD_LIST_NUMBER_STYLE_UPPER_LETTER,
            'lower_letter': self.WD_LIST_NUMBER_STYLE_LOWER_LETTER,
            'upper_roman': self.WD_LIST_NUMBER_STYLE_UPPER_ROMAN,
            'lower_roman': self.WD_LIST_NUMBER_STYLE_LOWER_ROMAN,
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
            return 'upper_letter'
        
        # Lower case letters: a., b., c., etc.
        if re.match(r'^[a-z]\.$', numbering):
            return 'lower_letter'
        
        # Roman numerals (basic detection)
        if re.match(r'^[IVX]+\.$', numbering):
            return 'upper_roman'
        
        # Default to decimal
        return 'decimal'
    
    def create_word_document(self, paragraphs: List[ParagraphData], output_path: str):
        """Create a new Word document with proper list levels"""
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
            
            # Track current list level
            current_level = 0
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
                    
                    # Apply numbering
                    self.apply_numbering_to_paragraph(para, numbering, level, list_objects, word)
                    
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
    
    def apply_numbering_to_paragraph(self, para, numbering: str, level: int, list_objects: Dict, word_app):
        """Apply numbering to a paragraph"""
        try:
            # Determine numbering style
            style = self.determine_numbering_style(numbering)
            
            # Get or create list object for this level
            if level not in list_objects:
                # Create new list
                list_obj = para.Range.ListFormat
                list_obj.ApplyListTemplate(
                    ListTemplate=word_app.ListGalleries(1).ListTemplates(1),
                    ContinuePreviousList=False,
                    ApplyTo=1  # wdListApplyToWholeList
                )
                list_objects[level] = list_obj
            else:
                # Use existing list
                para.Range.ListFormat = list_objects[level]
            
            # Set the list level
            para.Range.ListFormat.ListLevelNumber = level + 1
            
        except Exception as e:
            print(f"Warning: Could not apply numbering '{numbering}' at level {level}: {e}")
            # Fall back to plain text
            pass
    
    def reconstruct_document(self, json_path: str, output_path: str):
        """Main method to reconstruct a Word document from JSON analysis"""
        print(f"Loading JSON analysis from: {json_path}")
        json_data = self.load_json_analysis(json_path)
        
        print(f"Parsing {len(json_data.get('all_paragraphs', []))} paragraphs...")
        paragraphs = self.parse_paragraphs_from_json(json_data)
        
        print(f"Creating Word document with {len(paragraphs)} paragraphs...")
        self.create_word_document(paragraphs, output_path)
        
        print("Document reconstruction complete!")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python word_document_reconstructor.py <json_file> <output_docx>")
        print("Example: python word_document_reconstructor.py output/SECTION_00_00_00_hybrid_analysis.json reconstructed_SECTION_00_00_00.docx")
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
        reconstructor = WordDocumentReconstructor()
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