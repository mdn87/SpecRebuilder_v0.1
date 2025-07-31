#!/usr/bin/env python3
"""
Enhanced Hybrid Numbering Detector

This script combines Win32COM extraction with text-based deduction to find
all numbering patterns in Word documents, including those stored as text.
It also handles content blocks without numbering by appending them to previous blocks.
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
    cleaned_content: Optional[str] = None
    is_continuation: bool = False
    parent_index: Optional[int] = None

@dataclass
class ContentBlock:
    """Represents a consolidated content block with numbering"""
    index: int
    numbering: str  # True numbering or inferred numbering
    numbering_type: str  # "true", "inferred", or "none"
    content: str
    level: Optional[int] = None
    continuation_blocks: List[str] = None

@dataclass
class DocumentAnalysis:
    """Represents the analysis of a Word document"""
    total_paragraphs: int
    numbered_paragraphs: int
    unnumbered_paragraphs: int
    inferred_paragraphs: int
    continuation_paragraphs: int
    consolidated_blocks: int
    numbering_patterns: Dict[str, int]
    inferred_patterns: Dict[str, int]
    sample_paragraphs: List[Dict[str, Any]]
    content_blocks: List[ContentBlock]

class EnhancedHybridNumberingDetector:
    """Extracts and deduces numbering from Word documents using enhanced hybrid approach"""
    
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
                        # Clean the content by removing the inferred number and leading whitespace
                        paragraph.cleaned_content = self.clean_content_from_numbering(text, inferred_number)
                
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
    
    def clean_content_from_numbering(self, text: str, numbering: str) -> str:
        """Remove the inferred numbering and leading whitespace from text content"""
        if not text or not numbering:
            return text
        
        # Remove the numbering pattern from the beginning of the text
        # This handles cases like "3. \tGalvanized steel clamps..." -> "Galvanized steel clamps..."
        cleaned = text.strip()
        
        # Remove the numbering and any following whitespace/tabs
        if cleaned.startswith(numbering):
            cleaned = cleaned[len(numbering):].lstrip()
        
        return cleaned
    
    def consolidate_content_blocks(self, paragraphs: List[NumberedParagraph]) -> List[ContentBlock]:
        """Consolidate paragraphs into content blocks, appending unnumbered content to previous blocks"""
        content_blocks = []
        current_block = None
        continuation_texts = []
        
        for para in paragraphs:
            # Skip empty paragraphs
            if not para.text.strip():
                continue
            
            # Check if this paragraph has any numbering (true or inferred)
            has_numbering = bool(para.list_number or para.inferred_number)
            
            if has_numbering:
                # If we have a current block, save it with any continuation text
                if current_block:
                    if continuation_texts:
                        current_block.continuation_blocks = continuation_texts.copy()
                    content_blocks.append(current_block)
                
                # Start a new content block
                numbering = para.list_number if para.list_number else para.inferred_number
                numbering_type = "true" if para.list_number else "inferred"
                
                current_block = ContentBlock(
                    index=len(content_blocks),
                    numbering=numbering,
                    numbering_type=numbering_type,
                    content=para.text,
                    level=para.level,
                    continuation_blocks=[]
                )
                continuation_texts = []
                
            else:
                # This paragraph has no numbering - it's continuation content
                if current_block:
                    # Append to the current block
                    continuation_texts.append(para.text)
                else:
                    # No previous block to append to - create a standalone block
                    current_block = ContentBlock(
                        index=len(content_blocks),
                        numbering="",
                        numbering_type="none",
                        content=para.text,
                        level=None,
                        continuation_blocks=[]
                    )
                    continuation_texts = []
        
        # Don't forget the last block
        if current_block:
            if continuation_texts:
                current_block.continuation_blocks = continuation_texts.copy()
            content_blocks.append(current_block)
        
        return content_blocks
    
    def analyze_document_structure(self, paragraphs: List[NumberedParagraph]) -> DocumentAnalysis:
        """Analyze the structure of the extracted paragraphs"""
        
        numbered_count = sum(1 for p in paragraphs if p.list_number)
        inferred_count = sum(1 for p in paragraphs if p.inferred_number)
        continuation_count = sum(1 for p in paragraphs if not p.list_number and not p.inferred_number and p.text.strip())
        unnumbered_count = len(paragraphs) - numbered_count - inferred_count - continuation_count
        
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
        
        # Consolidate into content blocks
        content_blocks = self.consolidate_content_blocks(paragraphs)
        
        return DocumentAnalysis(
            total_paragraphs=len(paragraphs),
            numbered_paragraphs=numbered_count,
            unnumbered_paragraphs=unnumbered_count,
            inferred_paragraphs=inferred_count,
            continuation_paragraphs=continuation_count,
            consolidated_blocks=len(content_blocks),
            numbering_patterns=numbering_patterns,
            inferred_patterns=inferred_patterns,
            sample_paragraphs=sample_paragraphs,
            content_blocks=content_blocks
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
                'continuation_paragraphs': analysis.continuation_paragraphs,
                'unnumbered_paragraphs': analysis.unnumbered_paragraphs,
                'total_numbered': analysis.numbered_paragraphs + analysis.inferred_paragraphs,
                'consolidated_blocks': analysis.consolidated_blocks,
                'numbering_percentage': ((analysis.numbered_paragraphs + analysis.inferred_paragraphs) / analysis.total_paragraphs * 100) if analysis.total_paragraphs > 0 else 0
            },
            'numbering_patterns': analysis.numbering_patterns,
            'inferred_patterns': analysis.inferred_patterns,
            'sample_paragraphs': analysis.sample_paragraphs,
            'content_blocks': [
                {
                    'index': block.index,
                    'numbering': block.numbering,
                    'numbering_type': block.numbering_type,
                    'content': block.content,
                    'level': block.level,
                    'continuation_blocks': block.continuation_blocks
                }
                for block in analysis.content_blocks
            ],
            'all_paragraphs': [
                {
                    'index': para.index,
                    'list_number': para.list_number,
                    'inferred_number': para.inferred_number,
                    'text': para.text,
                    'combined': para.combined,
                    'cleaned_content': para.cleaned_content,
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
        content_blocks = report['content_blocks']
        
        print(f"\n=== ENHANCED HYBRID NUMBERING DETECTION SUMMARY ===")
        print(f"Document: {doc_info['filename']}")
        print(f"Path: {doc_info['path']}")
        print(f"Total paragraphs: {doc_info['total_paragraphs']}")
        
        print(f"\n=== STRUCTURE ANALYSIS ===")
        print(f"True numbered paragraphs: {structure['numbered_paragraphs']}")
        print(f"Inferred numbered paragraphs: {structure['inferred_paragraphs']}")
        print(f"Continuation paragraphs: {structure['continuation_paragraphs']}")
        print(f"Total numbered paragraphs: {structure['total_numbered']}")
        print(f"Unnumbered paragraphs: {structure['unnumbered_paragraphs']}")
        print(f"Consolidated content blocks: {structure['consolidated_blocks']}")
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
        
        print(f"\n=== CONSOLIDATED CONTENT BLOCKS ===")
        for i, block in enumerate(content_blocks[:15]):  # Show first 15 blocks
            numbering = block['numbering'] if block['numbering'] else "[No numbering]"
            numbering_type = block['numbering_type']
            continuation_count = len(block['continuation_blocks'])
            content_preview = block['content'][:60] + "..." if len(block['content']) > 60 else block['content']
            
            print(f"{i+1:2d}. [{numbering:8}] ({numbering_type}) '{content_preview}'")
            if continuation_count > 0:
                print(f"     + {continuation_count} continuation block(s)")
                for j, continuation in enumerate(block['continuation_blocks'][:2]):  # Show first 2 continuations
                    continuation_preview = continuation[:50] + "..." if len(continuation) > 50 else continuation
                    print(f"       {j+1}. '{continuation_preview}'")
                if continuation_count > 2:
                    print(f"       ... and {continuation_count - 2} more")
        
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
        print("Usage: python enhanced_hybrid_detector.py <docx_file> [output_dir]")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "output"
    
    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found: {docx_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        detector = EnhancedHybridNumberingDetector()
        
        print(f"Analyzing Word document with enhanced hybrid numbering detection: {docx_path}")
        
        # Generate analysis report
        report = detector.generate_analysis_report(docx_path)
        
        # Print summary
        detector.print_analysis_summary(report)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_enhanced_hybrid_analysis.json")
        detector.save_report(report, output_path)
        
        print(f"\nAnalysis report saved to: {output_path}")
        
    except ImportError as e:
        print(f"Error: {e}")
        print("Please install pywin32: pip install pywin32")
        sys.exit(1)
    except Exception as e:
        print(f"Error in enhanced hybrid numbering detection: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 