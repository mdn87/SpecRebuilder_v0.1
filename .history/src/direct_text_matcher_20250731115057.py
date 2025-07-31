#!/usr/bin/env python3
"""
Direct Text Matcher

This script extracts text directly from Word documents and matches it to
numbering patterns from text files using multiple strategies.
"""

import json
import sys
import os
import re
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from word_to_json import WordToJsonConverter

@dataclass
class TextExtraction:
    """Represents extracted text from Word document"""
    text: str
    index: int
    block_type: str = "content"
    raw_text: str = ""  # Original text before cleaning

@dataclass
class NumberingMatch:
    """Represents a match between numbering and extracted text"""
    numbering: str
    content: str
    extracted_text: str
    index: int
    confidence: float
    match_strategy: str
    level: Optional[int] = None

class DirectTextMatcher:
    """Extracts text directly from Word and matches to numbering"""
    
    def __init__(self):
        # MANUAL REFINEMENT ARRAYS
        
        # Text cleaning strategies
        self.text_cleaning_strategies = [
            "remove_extra_whitespace",
            "normalize_newlines",
            "remove_special_chars",
            "strip_punctuation",
            "lowercase_comparison"
        ]
        
        # Matching strategies (in order of preference)
        self.matching_strategies = [
            "exact_match",
            "contains_match",
            "fuzzy_match",
            "pattern_match",
            "word_overlap"
        ]
        
        # Confidence thresholds
        self.confidence_thresholds = {
            "exact_match": 1.0,
            "contains_match": 0.9,
            "fuzzy_match": 0.7,
            "pattern_match": 0.6,
            "word_overlap": 0.5
        }
        
        # Text extraction strategies
        self.extraction_strategies = [
            "paragraph_text",
            "run_text",
            "combined_text",
            "cleaned_text"
        ]
    
    def extract_text_from_word(self, docx_path: str) -> List[TextExtraction]:
        """Extract text directly from Word document using multiple strategies"""
        # Convert Word to JSON
        converter = WordToJsonConverter()
        json_path = converter.convert_to_json(docx_path)
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        extractions = []
        
        for paragraph in data.get('paragraphs', []):
            # Strategy 1: Direct paragraph text
            paragraph_text = paragraph.get('text', '').strip()
            
            # Strategy 2: Extract from runs (more detailed)
            runs_text = ""
            for run in paragraph.get('runs', []):
                runs_text += run.get('text', '')
            runs_text = runs_text.strip()
            
            # Strategy 3: Combined approach
            combined_text = paragraph_text if paragraph_text else runs_text
            
            # Strategy 4: Cleaned text
            cleaned_text = self.clean_text(combined_text)
            
            if combined_text:  # Only include non-empty paragraphs
                extraction = TextExtraction(
                    text=cleaned_text,
                    index=len(extractions),
                    raw_text=combined_text
                )
                extractions.append(extraction)
        
        # Classify blocks based on position
        if len(extractions) >= 3:
            extractions[0].block_type = "section_number"
            extractions[1].block_type = "section_title"
            extractions[-1].block_type = "end_of_section"
            for i in range(2, len(extractions) - 1):
                extractions[i].block_type = "content"
        
        return extractions
    
    def clean_text(self, text: str) -> str:
        """Clean text for better matching"""
        # Remove extra whitespace
        text = re.sub(r'\s+', ' ', text)
        
        # Normalize newlines
        text = text.replace('\n', ' ').replace('\r', ' ')
        
        # Remove special characters that might interfere with matching
        text = re.sub(r'[^\w\s\-\.]', '', text)
        
        # Strip leading/trailing whitespace
        text = text.strip()
        
        return text
    
    def extract_numbering_from_text(self, txt_path: str) -> List[Dict[str, Any]]:
        """Extract numbering patterns from text file"""
        numbered_lines = []
        
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for i, line in enumerate(lines, 1):
            line = line.strip()
            if not line or line in ['SECTION 00 00 00', 'SECTION TITLE', 'END OF SECTION']:
                continue
            
            # Try to find numbering at start of line
            numbering, content = self.extract_numbering_from_line(line)
            
            if numbering and content:
                numbered_lines.append({
                    'line_number': i,
                    'numbering': numbering,
                    'content': content,
                    'level': self.determine_level(numbering)
                })
        
        return numbered_lines
    
    def extract_numbering_from_line(self, line: str) -> Tuple[Optional[str], Optional[str]]:
        """Extract numbering from a single line"""
        # Try different separator patterns
        separators = ['\t', '  ', ' - ', ' -', '- ', ' )', ') ', ' ]', '] ']
        
        for sep in separators:
            if sep in line:
                parts = line.split(sep, 1)
                if len(parts) >= 2:
                    potential_numbering = parts[0].strip()
                    potential_content = parts[1].strip()
                    
                    if self.looks_like_numbering(potential_numbering):
                        return potential_numbering, potential_content
        
        # Fallback: try to find numbering at start
        # Look for patterns like "1.0", "A.", "1.", "a.", "i."
        patterns = [
            r'^(\d+\.0)\s+(.+)',
            r'^(\d+\.\d+)\s+(.+)',
            r'^([A-Z]\.)\s+(.+)',
            r'^(\d+\.)\s+(.+)',
            r'^([a-z]\.)\s+(.+)',
            r'^([ivxlcdm]+\.)\s+(.+)',
        ]
        
        for pattern in patterns:
            match = re.match(pattern, line, re.IGNORECASE)
            if match:
                return match.group(1), match.group(2)
        
        return None, None
    
    def looks_like_numbering(self, text: str) -> bool:
        """Check if text looks like numbering"""
        # Remove punctuation for analysis
        clean_text = re.sub(r'[^\w]', '', text)
        
        # Check for common numbering patterns
        patterns = [
            r'^\d+$',  # Just numbers
            r'^[A-Z]+$',  # Just uppercase letters
            r'^[a-z]+$',  # Just lowercase letters
            r'^[ivxlcdm]+$',  # Roman numerals
            r'^\d+[A-Z]+$',  # Numbers + letters
            r'^[A-Z]+\d+$',  # Letters + numbers
        ]
        
        for pattern in patterns:
            if re.match(pattern, clean_text, re.IGNORECASE):
                return True
        
        return False
    
    def determine_level(self, numbering: str) -> Optional[int]:
        """Determine level from numbering"""
        # Remove punctuation for analysis
        clean_num = re.sub(r'[^\w]', '', numbering)
        
        if re.match(r'^\d+0$', clean_num):  # 10, 20, etc.
            return 0
        elif re.match(r'^\d+\d+$', clean_num):  # 11, 12, 21, etc.
            return 1
        elif re.match(r'^[A-Z]+$', clean_num):  # A, B, C, etc.
            return 2
        elif re.match(r'^\d+$', clean_num):  # 1, 2, 3, etc.
            return 3
        elif re.match(r'^[a-z]+$', clean_num):  # a, b, c, etc.
            return 4
        elif re.match(r'^[ivxlcdm]+$', clean_num, re.IGNORECASE):  # i, ii, iii, etc.
            return 5
        
        return None
    
    def match_numbering_to_text(self, numbered_lines: List[Dict[str, Any]], text_extractions: List[TextExtraction]) -> List[NumberingMatch]:
        """Match numbering to extracted text using multiple strategies"""
        matches = []
        
        for numbered_line in numbered_lines:
            best_match = None
            best_confidence = 0.0
            best_strategy = ""
            
            for extraction in text_extractions:
                if extraction.block_type != "content":
                    continue
                
                # Try different matching strategies
                for strategy in self.matching_strategies:
                    confidence = self.calculate_match_confidence(
                        numbered_line['content'], 
                        extraction.text, 
                        strategy
                    )
                    
                    if confidence > best_confidence:
                        best_confidence = confidence
                        best_match = NumberingMatch(
                            numbering=numbered_line['numbering'],
                            content=numbered_line['content'],
                            extracted_text=extraction.text,
                            index=extraction.index,
                            confidence=confidence,
                            match_strategy=strategy,
                            level=numbered_line.get('level')
                        )
                        best_strategy = strategy
            
            if best_match and best_match.confidence > 0.3:  # Minimum threshold
                matches.append(best_match)
        
        return matches
    
    def calculate_match_confidence(self, expected_content: str, extracted_text: str, strategy: str) -> float:
        """Calculate confidence for a match using specified strategy"""
        
        if strategy == "exact_match":
            # Exact text match (case insensitive)
            if expected_content.lower() == extracted_text.lower():
                return 1.0
        
        elif strategy == "contains_match":
            # One contains the other
            if expected_content.lower() in extracted_text.lower():
                return 0.9
            elif extracted_text.lower() in expected_content.lower():
                return 0.8
        
        elif strategy == "fuzzy_match":
            # Fuzzy matching using similarity
            similarity = self.calculate_text_similarity(expected_content, extracted_text)
            return similarity * 0.7
        
        elif strategy == "pattern_match":
            # Pattern-based matching (e.g., BWA- patterns)
            if "BWA-" in expected_content and "BWA-" in extracted_text:
                return 0.6
        
        elif strategy == "word_overlap":
            # Word overlap matching
            overlap_score = self.calculate_word_overlap(expected_content, extracted_text)
            return overlap_score * 0.5
        
        return 0.0
    
    def calculate_text_similarity(self, text1: str, text2: str) -> float:
        """Calculate similarity between two texts"""
        # Simple character-based similarity
        if not text1 or not text2:
            return 0.0
        
        # Convert to sets of characters
        chars1 = set(text1.lower())
        chars2 = set(text2.lower())
        
        if not chars1 or not chars2:
            return 0.0
        
        intersection = chars1.intersection(chars2)
        union = chars1.union(chars2)
        
        return len(intersection) / len(union)
    
    def calculate_word_overlap(self, text1: str, text2: str) -> float:
        """Calculate word overlap between two texts"""
        words1 = set(text1.lower().split())
        words2 = set(text2.lower().split())
        
        if not words1 or not words2:
            return 0.0
        
        intersection = words1.intersection(words2)
        union = words1.union(words2)
        
        return len(intersection) / len(union)
    
    def generate_direct_matching_report(self, docx_path: str, txt_path: str) -> Dict[str, Any]:
        """Generate a comprehensive direct text matching report"""
        
        # Extract text from Word document
        print("Extracting text from Word document...")
        text_extractions = self.extract_text_from_word(docx_path)
        print(f"Found {len(text_extractions)} text extractions")
        
        # Extract numbering from text file
        print("Extracting numbering from text file...")
        numbered_lines = self.extract_numbering_from_text(txt_path)
        print(f"Found {len(numbered_lines)} numbered lines")
        
        # Match numbering to text
        print("Matching numbering to extracted text...")
        matches = self.match_numbering_to_text(numbered_lines, text_extractions)
        print(f"Found {len(matches)} matches")
        
        # Generate report
        report = {
            'text_extractions': [
                {
                    'index': extraction.index,
                    'text': extraction.text,
                    'raw_text': extraction.raw_text,
                    'block_type': extraction.block_type
                }
                for extraction in text_extractions
            ],
            'numbered_lines': numbered_lines,
            'matches': [
                {
                    'numbering': match.numbering,
                    'content': match.content,
                    'extracted_text': match.extracted_text,
                    'index': match.index,
                    'confidence': match.confidence,
                    'match_strategy': match.match_strategy,
                    'level': match.level
                }
                for match in matches
            ],
            'summary': {
                'total_extractions': len(text_extractions),
                'total_numbered_lines': len(numbered_lines),
                'total_matches': len(matches),
                'match_rate': len(matches) / len(numbered_lines) if numbered_lines else 0,
                'average_confidence': sum(match.confidence for match in matches) / len(matches) if matches else 0
            }
        }
        
        return report
    
    def print_direct_matching_summary(self, report: Dict[str, Any]):
        """Print a summary of the direct matching results"""
        summary = report['summary']
        
        print(f"\n=== DIRECT TEXT MATCHING SUMMARY ===")
        print(f"Text extractions from Word: {summary['total_extractions']}")
        print(f"Numbered lines from text: {summary['total_numbered_lines']}")
        print(f"Successful matches: {summary['total_matches']}")
        print(f"Match rate: {summary['match_rate']:.2%}")
        print(f"Average confidence: {summary['average_confidence']:.2f}")
        
        print(f"\n=== TOP MATCHES ===")
        matches = sorted(report['matches'], key=lambda x: x['confidence'], reverse=True)
        for i, match in enumerate(matches[:10]):
            print(f"{i+1:2d}. [{match['confidence']:.2f}] '{match['numbering']}' → '{match['extracted_text'][:50]}'")
        
        print(f"\n=== SAMPLE TEXT EXTRACTIONS ===")
        extractions = report['text_extractions']
        for i, extraction in enumerate(extractions[:10]):
            print(f"{i+1:2d}. [{extraction['block_type']:15}] '{extraction['text'][:50]}'")
        
        print(f"\n=== UNMATCHED NUMBERED LINES ===")
        matched_numberings = {match['numbering'] for match in matches}
        unmatched_lines = [line for line in report['numbered_lines'] if line['numbering'] not in matched_numberings]
        for line in unmatched_lines[:10]:
            print(f"'{line['numbering']}' → '{line['content']}'")
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save the direct matching report"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"Direct matching report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python direct_text_matcher.py <docx_file> <txt_file> [output_dir]")
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
    
    matcher = DirectTextMatcher()
    
    try:
        print(f"Extracting text directly from: {docx_path}")
        print(f"Matching to numbering from: {txt_path}")
        
        # Generate direct matching report
        report = matcher.generate_direct_matching_report(docx_path, txt_path)
        
        # Print summary
        matcher.print_direct_matching_summary(report)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_direct_text_matches.json")
        matcher.save_report(report, output_path)
        
        print(f"\nDirect matching report saved to: {output_path}")
        
    except Exception as e:
        print(f"Error in direct text matching: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 