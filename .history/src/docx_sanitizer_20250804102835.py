#!/usr/bin/env python3
"""
Word Document Sanitizer

This script uses python-docx to load and resave Word documents,
which normalizes the internal structure and eliminates "unreadable content" warnings.
"""

import os
import sys
from docx import Document

def sanitize_docx(input_path: str, output_path: str):
    """Load and resave a Word document to normalize its structure"""
    try:
        print(f"Loading document: {input_path}")
        doc = Document(input_path)
        
        print(f"Saving sanitized document: {output_path}")
        doc.save(output_path)
        
        print("Document sanitization complete!")
        return True
        
    except Exception as e:
        print(f"Error sanitizing document: {e}")
        return False

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python docx_sanitizer.py <input_docx> <output_docx>")
        print("Example: python docx_sanitizer.py output/word_compatible_output.docx output/word_compatible_output_cleaned.docx")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(input_path):
        print(f"Error: Input document not found: {input_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    success = sanitize_docx(input_path, output_path)
    if not success:
        sys.exit(1)

if __name__ == "__main__":
    main() 