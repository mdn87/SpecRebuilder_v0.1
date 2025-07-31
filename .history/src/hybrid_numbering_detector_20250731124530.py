#!/usr/bin/env python3
"""
Hybrid Numbering Detector

This script combines Win32COM extraction with text-based deduction to find
all numbering patterns in Word documents, including those stored as text.
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
    level: Optional[int] = None
    inferred_number: Optional[str] = None
    deduction_method: Optional[str] = None

@dataclass
class DocumentAnalysis:
    """Represents the analysis of a Word document"""
    total_paragraphs: int
    numbered_paragraphs: int
    unnumbered_paragraphs: int
    inferred_paragraphs: int
    numbering_patterns: Dict[str, int]
    inferred_patterns: Dict[str, int]
    sample_paragraphs: List[Dict[str, Any]]

class HybridNumberingDetector:
    """Extracts and deduces numbering from Word documents using hybrid approach"""
    
    def __init__(self):
        if not WIN32COM_AVAILABLE:
            raise ImportError("win32com not available. Install with: pip install pywin32")
        
        # Word constant meaning "no automatic numbering"
        self.WD_LIST_NO_NUMBERING = 0
        
        # Common numbering patterns to look for in text
        self.numbering_patterns = [
            r'^(\d+\.\d+)\s*',  # 1.0, 1.01, 2.0, etc.
            r'^(\d+\.)\s*',     # 1., 2., 3., etc.
            r'^([A-Z]\.)\s*',   # A., B., C., etc.
            r'^([a-z]\.)\s*',   # a., b., c., etc.
            r'^\((\d+\))\s*',   # (1), (2), (3), etc.
            r'^\(([A-Z]\))\s*', # (A), (B), (C), etc.
            r'^\(([a-z]\))\s*', # (a), (b), (c), etc.
        ]
    
    def extract_numbered_paragraphs(self, doc_path: str) -> List[NumberedParagraph]:
        """
        Returns list of NumberedParagraph objects in document order,
        with both true numbering and inferred numbering.
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
                level = None
                
                try:
                    lf = para.Range.ListFormat
                    if lf.ListType != self.WD_LIST_NO_NUMBERING:
                        # This is the numbering string Word would show, e.g., "1.01", "a.", etc.
                        list_number = lf.ListString or ""
                        # Try to get the level
                        try:
                            level = lf.ListLevelNumber
                        except:
                            level = None
                except Exception:
                    list_number = ""
                
                if list_number:
                    combined = f"{list_number}\t{text}"
                else:
                    combined = text
                
                # Create the paragraph object
                paragraph = NumberedParagraph(
                    index=idx,
                    list_number=list_number,
                    text=text,
                    combined=combined,
                    level=level
                )
                
                # If no true numbering was found, try to deduce it
                if not list_number and text.strip():
                    inferred_number = self.deduce_numbering_from_text(text)
                    if inferred_number:
                        paragraph.inferred_number = inferred_number
                        paragraph.deduction_method = "text_pattern"
                
                results.append(paragraph)
                
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
    
    def deduce_numbering_from_text(self, text: str) -> Optional[str]:
        """Deduce numbering from text content using pattern matching"""
        if not text or not text.strip():
            return None
        
        # Try each numbering pattern
        for pattern in self.numbering_patterns:
            match = re.match(pattern, text.strip())
            if match:
                return match.group(1)
        
        return None
    
    def analyze_document_structure(self, paragraphs: List[NumberedParagraph]) -> DocumentAnalysis:
        """Analyze the structure of the extracted paragraphs"""
        
        numbered_count = sum(1 for p in paragraphs if p.list_number)
        inferred_count = sum(1 for p in paragraphs if p.inferred_number)
        unnumbered_count = len(paragraphs) - numbered_count - inferred_count
        
        # Count numbering patterns
        numbering_patterns = {}
        for para in paragraphs:
            if para.list_number:
                numbering_patterns[para.list_number] = numbering_patterns.get(para.list_number, 0) + 1
        
        # Count inferred patterns
        inferred_patterns = {}
        for para in paragraphs:
            if para.inferred_number:
                inferred_patterns[para.inferred_number] = inferred_patterns.get(para.inferred_number, 0) + 1
        
        # Get sample paragraphs
        sample_paragraphs = []
        for para in paragraphs[:20]:  # First 20 paragraphs
            sample_paragraphs.append({
                'index': para.index,
                'list_number': para.list_number,
                'inferred_number': para.inferred_number,
                'text': para.text[:100] + "..." if len(para.text) > 100 else para.text,
                'combined': para.combined[:100] + "..." if len(para.combined) > 100 else para.combined,
                'level': para.level,
                'deduction_method': para.deduction_method
            })
        
        return DocumentAnalysis(
            total_paragraphs=len(paragraphs),
            numbered_paragraphs=numbered_count,
            unnumbered_paragraphs=unnumbered_count,
            inferred_paragraphs=inferred_count,
            numbering_patterns=numbering_patterns,
            inferred_patterns=inferred_patterns,
            sample_paragraphs=sample_paragraphs
        )
    
    def generate_analysis_report(self, docx_path: str) -> Dict[str, Any]:
        """Generate a comprehensive analysis report"""
        
        # Extract numbered paragraphs from Word document
        print("Extracting numbered paragraphs from Word document...")
        paragraphs = self.extract_numbered_paragraphs(docx_path)
        print(f"Found {len(paragraphs)} paragraphs in Word document")
        
        # Analyze document structure
        print("Analyzing document structure...")
        analysis = self.analyze_document_structure(paragraphs)
        
        # Generate report
        report = {
            'document_info': {
                'path': docx_path,
                'filename': Path(docx_path).name,
                'total_paragraphs': analysis.total_paragraphs
            },
            'structure_analysis': {
                'numbered_paragraphs': analysis.numbered_paragraphs,
                'inferred_paragraphs': analysis.inferred_paragraphs,
                'unnumbered_paragraphs': analysis.unnumbered_paragraphs,
                'total_numbered': analysis.numbered_paragraphs + analysis.inferred_paragraphs,
                'numbering_percentage': ((analysis.numbered_paragraphs + analysis.inferred_paragraphs) / analysis.total_paragraphs * 100) if analysis.total_paragraphs > 0 else 0
            },
            'numbering_patterns': analysis.numbering_patterns,
            'inferred_patterns': analysis.inferred_patterns,
            'sample_paragraphs': analysis.sample_paragraphs,
            'all_paragraphs': [
                {
                    'index': para.index,
                    'list_number': para.list_number,
                    'inferred_number': para.inferred_number,
                    'text': para.text,
                    'combined': para.combined,
                    'level': para.level,
                    'deduction_method': para.deduction_method
                }
                for para in paragraphs
            ]
        }
        
        return report
    
    def print_analysis_summary(self, report: Dict[str, Any]):
        """Print a summary of the analysis results"""
        doc_info = report['document_info']
        structure = report['structure_analysis']
        patterns = report['numbering_patterns']
        inferred_patterns = report['inferred_patterns']
        
        print(f"\n=== HYBRID NUMBERING DETECTION SUMMARY ===")
        print(f"Document: {doc_info['filename']}")
        print(f"Path: {doc_info['path']}")
        print(f"Total paragraphs: {doc_info['total_paragraphs']}")
        
        print(f"\n=== STRUCTURE ANALYSIS ===")
        print(f"True numbered paragraphs: {structure['numbered_paragraphs']}")
        print(f"Inferred numbered paragraphs: {structure['inferred_paragraphs']}")
        print(f"Total numbered paragraphs: {structure['total_numbered']}")
        print(f"Unnumbered paragraphs: {structure['unnumbered_paragraphs']}")
        print(f"Overall numbering percentage: {structure['numbering_percentage']:.1f}%")
        
        print(f"\n=== TRUE NUMBERING PATTERNS ===")
        if patterns:
            sorted_patterns = sorted(patterns.items(), key=lambda x: x[1], reverse=True)
            for pattern, count in sorted_patterns[:10]:  # Top 10 patterns
                print(f"  '{pattern}': {count} occurrences")
        else:
            print("  No true numbering patterns found")
        
        print(f"\n=== INFERRED NUMBERING PATTERNS ===")
        if inferred_patterns:
            sorted_inferred = sorted(inferred_patterns.items(), key=lambda x: x[1], reverse=True)
            for pattern, count in sorted_inferred[:10]:  # Top 10 patterns
                print(f"  '{pattern}': {count} occurrences")
        else:
            print("  No inferred numbering patterns found")
        
        print(f"\n=== SAMPLE PARAGRAPHS ===")
        for i, para in enumerate(report['sample_paragraphs']):
            numbering = para['list_number'] if para['list_number'] else "[No true numbering]"
            inferred = para['inferred_number'] if para['inferred_number'] else "[No inferred numbering]"
            level_info = f" (Level {para['level']})" if para['level'] is not None else ""
            method_info = f" [{para['deduction_method']}]" if para['deduction_method'] else ""
            print(f"{i+1:2d}. True: [{numbering:8}] Inferred: [{inferred:8}]{level_info}{method_info} '{para['text']}'")
        
        # Assess the document
        numbering_percentage = structure['numbering_percentage']
        if numbering_percentage >= 50:
            print(f"\n✅ EXCELLENT: {numbering_percentage:.1f}% of paragraphs have numbering - Document has good structure!")
        elif numbering_percentage >= 25:
            print(f"\n⚠️  GOOD: {numbering_percentage:.1f}% of paragraphs have numbering - Document has some structure.")
        else:
            print(f"\n❌ POOR: {numbering_percentage:.1f}% of paragraphs have numbering - Document lacks structure.")
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save the analysis report"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"Analysis report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 2:
        print("Usage: python hybrid_numbering_detector.py <docx_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    
    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found: {docx_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        detector = HybridNumberingDetector()
        
        print(f"Analyzing Word document with hybrid numbering detection: {docx_path}")
        
        # Generate analysis report
        report = detector.generate_analysis_report(docx_path)
        
        # Print summary
        detector.print_analysis_summary(report)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_hybrid_analysis.json")
        detector.save_report(report, output_path)
        
        print(f"\nAnalysis report saved to: {output_path}")
        
    except ImportError as e:
        print(f"Error: {e}")
        print("Please install pywin32: pip install pywin32")
        sys.exit(1)
    except Exception as e:
        print(f"Error in hybrid numbering detection: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 