#!/usr/bin/env python3
"""
Win32COM Word Extractor

This script uses win32com.client to extract numbered paragraphs directly from Word
documents and compare them to text files, using the approach from the instructions.
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
class NumberedParagraph:
    """Represents a numbered paragraph from Word"""
    index: int
    list_number: str
    text: str
    combined: str

@dataclass
class ComparisonResult:
    """Represents a comparison between Word and text file"""
    word_paragraph: NumberedParagraph
    text_file_line: str
    is_exact_match: bool
    differences: List[str]

class Win32COMExtractor:
    """Extracts numbered paragraphs from Word documents using win32com"""
    
    def __init__(self):
        if not WIN32COM_AVAILABLE:
            raise ImportError("win32com not available. Install with: pip install pywin32")
        
        # Word constant meaning "no automatic numbering"
        self.WD_LIST_NO_NUMBERING = 0
    
    def extract_numbered_paragraphs(self, doc_path: str) -> List[NumberedParagraph]:
        """
        Returns list of NumberedParagraph objects in document order,
        matching what Word would put when copying into Notepad (e.g., "1.01\tTitle").
        """
        # Necessary if called from a thread; safe to call multiple times per thread
        pythoncom.CoInitialize()
        word = None
        doc = None
        results = []
        
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False  # headless
            # Open readonly, don't prompt
            doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True, AddToRecentFiles=False)
            
            for idx, para in enumerate(doc.Paragraphs):
                raw = para.Range.Text or ""
                # Strip trailing paragraph mark(s)
                text = raw.rstrip("\r\x07")
                list_number = ""
                
                try:
                    lf = para.Range.ListFormat
                    if lf.ListType != self.WD_LIST_NO_NUMBERING:
                        # This is the numbering string Word would show, e.g., "1.01", "a.", etc.
                        list_number = lf.ListString or ""
                except Exception:
                    list_number = ""
                
                if list_number:
                    combined = f"{list_number}\t{text}"
                else:
                    combined = text
                
                results.append(NumberedParagraph(
                    index=idx,
                    list_number=list_number,
                    text=text,
                    combined=combined
                ))
                
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            # Clean up COM apartment if you initialized it manually
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        
        return results
    
    def read_text_file(self, txt_path: str) -> List[str]:
        """Read all lines from text file"""
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Strip whitespace and filter out empty lines
        text_lines = []
        for line in lines:
            stripped_line = line.strip()
            if stripped_line:  # Only include non-empty lines
                text_lines.append(stripped_line)
        
        return text_lines
    
    def compare_word_to_text(self, word_paragraphs: List[NumberedParagraph], text_lines: List[str]) -> List[ComparisonResult]:
        """Compare Word paragraphs with text file lines"""
        comparisons = []
        
        # Use the shorter list length to avoid index errors
        max_lines = min(len(word_paragraphs), len(text_lines))
        
        for i in range(max_lines):
            word_para = word_paragraphs[i]
            text_line = text_lines[i]
            
            # Compare the combined string (which includes numbering) with text file line
            is_exact_match = word_para.combined == text_line
            differences = []
            
            if not is_exact_match:
                # Check if it's just whitespace differences
                if word_para.combined.strip() == text_line.strip():
                    differences.append("Whitespace differences only")
                else:
                    differences.append(f"Word: '{word_para.combined[:50]}...' vs Text: '{text_line[:50]}...'")
            
            comparison = ComparisonResult(
                word_paragraph=word_para,
                text_file_line=text_line,
                is_exact_match=is_exact_match,
                differences=differences
            )
            comparisons.append(comparison)
        
        # Handle any remaining lines
        if len(word_paragraphs) > len(text_lines):
            for i in range(len(text_lines), len(word_paragraphs)):
                comparison = ComparisonResult(
                    word_paragraph=word_paragraphs[i],
                    text_file_line="",
                    is_exact_match=False,
                    differences=["Extra paragraph in Word document"]
                )
                comparisons.append(comparison)
        
        elif len(text_lines) > len(word_paragraphs):
            for i in range(len(word_paragraphs), len(text_lines)):
                comparison = ComparisonResult(
                    word_paragraph=NumberedParagraph(index=i, list_number="", text="", combined=""),
                    text_file_line=text_lines[i],
                    is_exact_match=False,
                    differences=["Extra line in text file"]
                )
                comparisons.append(comparison)
        
        return comparisons
    
    def generate_comparison_report(self, docx_path: str, txt_path: str) -> Dict[str, Any]:
        """Generate a comprehensive comparison report"""
        
        # Extract numbered paragraphs from Word document
        print("Extracting numbered paragraphs from Word document...")
        word_paragraphs = self.extract_numbered_paragraphs(docx_path)
        print(f"Found {len(word_paragraphs)} paragraphs in Word document")
        
        # Read text file
        print("Reading text file...")
        text_lines = self.read_text_file(txt_path)
        print(f"Found {len(text_lines)} lines in text file")
        
        # Compare Word paragraphs with text file
        print("Comparing Word paragraphs with text file...")
        comparisons = self.compare_word_to_text(word_paragraphs, text_lines)
        print(f"Generated {len(comparisons)} comparisons")
        
        # Calculate statistics
        exact_matches = sum(1 for comp in comparisons if comp.is_exact_match)
        total_comparisons = len(comparisons)
        match_percentage = exact_matches / total_comparisons if total_comparisons > 0 else 0
        
        # Generate report
        report = {
            'word_document': {
                'path': docx_path,
                'total_paragraphs': len(word_paragraphs),
                'sample_paragraphs': [
                    {
                        'index': para.index,
                        'list_number': para.list_number,
                        'text': para.text,
                        'combined': para.combined
                    }
                    for para in word_paragraphs[:10]
                ]
            },
            'text_file': {
                'path': txt_path,
                'total_lines': len(text_lines),
                'sample_lines': text_lines[:10]
            },
            'comparison_results': {
                'total_comparisons': total_comparisons,
                'exact_matches': exact_matches,
                'mismatches': total_comparisons - exact_matches,
                'match_percentage': match_percentage
            },
            'detailed_comparisons': [
                {
                    'index': comp.word_paragraph.index,
                    'word_combined': comp.word_paragraph.combined,
                    'text_file_line': comp.text_file_line,
                    'is_exact_match': comp.is_exact_match,
                    'differences': comp.differences
                }
                for comp in comparisons
            ]
        }
        
        return report
    
    def print_comparison_summary(self, report: Dict[str, Any]):
        """Print a summary of the comparison results"""
        word_doc = report['word_document']
        text_file = report['text_file']
        results = report['comparison_results']
        
        print(f"\n=== WIN32COM EXTRACTION SUMMARY ===")
        print(f"Word document: {word_doc['path']}")
        print(f"  - Total paragraphs: {word_doc['total_paragraphs']}")
        print(f"Text file: {text_file['path']}")
        print(f"  - Total lines: {text_file['total_lines']}")
        
        print(f"\n=== COMPARISON RESULTS ===")
        print(f"Total comparisons: {results['total_comparisons']}")
        print(f"Exact matches: {results['exact_matches']}")
        print(f"Mismatches: {results['mismatches']}")
        print(f"Match percentage: {results['match_percentage']:.2%}")
        
        print(f"\n=== SAMPLE WORD PARAGRAPHS ===")
        for i, para in enumerate(word_doc['sample_paragraphs']):
            numbering = para['list_number'] if para['list_number'] else "[No numbering]"
            print(f"{i+1:2d}. [{numbering:8}] '{para['text'][:50]}...'")
        
        print(f"\n=== SAMPLE TEXT FILE LINES ===")
        for i, line in enumerate(text_file['sample_lines']):
            print(f"{i+1:2d}. '{line[:50]}...'")
        
        print(f"\n=== SAMPLE COMPARISONS ===")
        comparisons = report['detailed_comparisons']
        for i, comp in enumerate(comparisons[:10]):
            status = "✓" if comp['is_exact_match'] else "✗"
            print(f"{i+1:2d}. [{status}] Word: '{comp['word_combined'][:50]}...'")
        
        print(f"\n=== MISMATCHES ===")
        mismatches = [comp for comp in comparisons if not comp['is_exact_match']]
        for i, comp in enumerate(mismatches[:10]):
            print(f"{i+1:2d}. {comp['differences']}")
        
        if results['match_percentage'] >= 0.95:
            print(f"\n✅ EXCELLENT: {results['match_percentage']:.1%} match rate - Win32COM extraction is highly accurate!")
        elif results['match_percentage'] >= 0.80:
            print(f"\n⚠️  GOOD: {results['match_percentage']:.1%} match rate - Win32COM extraction is mostly accurate.")
        else:
            print(f"\n❌ POOR: {results['match_percentage']:.1%} match rate - Win32COM extraction needs improvement.")
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save the comparison report"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"Comparison report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python win32com_extractor.py <docx_file> <txt_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    txt_path = sys.argv[2]
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "output"
    
    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found: {docx_path}")
        sys.exit(1)
    
    if not os.path.exists(txt_path):
        print(f"Error: TXT file not found: {txt_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        extractor = Win32COMExtractor()
        
        print(f"Extracting numbered paragraphs from: {docx_path}")
        print(f"Comparing with text file: {txt_path}")
        
        # Generate comparison report
        report = extractor.generate_comparison_report(docx_path, txt_path)
        
        # Print summary
        extractor.print_comparison_summary(report)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_win32com_comparison.json")
        extractor.save_report(report, output_path)
        
        print(f"\nComparison report saved to: {output_path}")
        
    except ImportError as e:
        print(f"Error: {e}")
        print("Please install pywin32: pip install pywin32")
        sys.exit(1)
    except Exception as e:
        print(f"Error in Win32COM extraction: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 