#!/usr/bin/env python3
"""
Numbering Pattern Matcher

This script detects numbering patterns from a text file and matches them to
content blocks from a Word document. It uses arrays for manual refinement.
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
class NumberingPattern:
    """Represents a numbering pattern found in text"""
    pattern: str
    level: int
    description: str

@dataclass
class TextLineMatch:
    """Represents a line from text file with numbering"""
    line_number: int
    numbering: str
    content: str
    pattern_matched: str
    level: Optional[int] = None

@dataclass
class ContentBlock:
    """Represents a content block from Word document"""
    text: str
    index: int
    block_type: str = "content"

@dataclass
class NumberingMatch:
    """Represents a match between text numbering and content block"""
    text_line: TextLineMatch
    content_block: ContentBlock
    confidence: float
    match_type: str

class NumberingPatternMatcher:
    """Matches numbering patterns from text to Word document content"""
    
    def __init__(self):
        # MANUAL REFINEMENT ARRAYS - Edit these as needed
        
        # Numbering patterns to detect (regex patterns)
        self.numbering_patterns = [
            NumberingPattern(r'^\d+\.0\s+', 0, "Major section (1.0, 2.0)"),
            NumberingPattern(r'^\d+\.\d+\s+', 1, "Subsection (1.01, 1.02)"),
            NumberingPattern(r'^[A-Z]\.\s+', 2, "Upper letter (A., B.)"),
            NumberingPattern(r'^\d+\.\s+', 3, "Decimal (1., 2.)"),
            NumberingPattern(r'^[a-z]\.\s+', 4, "Lower letter (a., b.)"),
            NumberingPattern(r'^[ivxlcdm]+\.\s+', 5, "Roman numeral (i., ii.)"),
        ]
        
        # Separator characters that indicate numbering
        self.separator_chars = ['\t', ' ', '.', '-', ')', ']']
        
        # Content matching strategies (in order of preference)
        self.matching_strategies = [
            "exact_text_match",
            "contains_text_match", 
            "fuzzy_text_match",
            "pattern_based_match"
        ]
        
        # Confidence thresholds
        self.confidence_thresholds = {
            "exact_text_match": 1.0,
            "contains_text_match": 0.8,
            "fuzzy_text_match": 0.6,
            "pattern_based_match": 0.4
        }
    
    def extract_content_blocks_from_word(self, docx_path: str) -> List[ContentBlock]:
        """Extract content blocks from Word document (removing blank lines)"""
        # Convert Word to JSON
        converter = WordToJsonConverter()
        json_path = converter.convert_to_json(docx_path)
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        content_blocks = []
        for paragraph in data.get('paragraphs', []):
            text = paragraph.get('text', '').strip()
            if text:  # Only include non-empty paragraphs
                block = ContentBlock(
                    text=text,
                    index=len(content_blocks)
                )
                content_blocks.append(block)
        
        # Classify blocks based on position
        if len(content_blocks) >= 3:
            # First block: section number
            content_blocks[0].block_type = "section_number"
            # Second block: section title  
            content_blocks[1].block_type = "section_title"
            # Last block: end of section
            content_blocks[-1].block_type = "end_of_section"
            # All others are content blocks
            for i in range(2, len(content_blocks) - 1):
                content_blocks[i].block_type = "content"
        
        return content_blocks
    
    def extract_numbering_from_text(self, txt_path: str) -> List[TextLineMatch]:
        """Extract numbering patterns from text file"""
        text_lines = []
        
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for i, line in enumerate(lines, 1):
            line = line.strip()
            if not line or line in ['SECTION 00 00 00', 'SECTION TITLE', 'END OF SECTION']:
                continue
            
            # Try to match numbering patterns
            numbering = None
            content = None
            pattern_matched = None
            level = None
            
            for pattern in self.numbering_patterns:
                match = re.match(pattern.pattern, line)
                if match:
                    numbering = match.group(0).strip()
                    content = line[len(match.group(0)):].strip()
                    pattern_matched = pattern.pattern
                    level = pattern.level
                    break
            
            # If no pattern matched, try to find numbering manually
            if not numbering:
                # Look for common numbering patterns at start of line
                for sep in self.separator_chars:
                    parts = line.split(sep, 1)
                    if len(parts) >= 2:
                        potential_numbering = parts[0].strip()
                        potential_content = parts[1].strip()
                        
                        # Check if potential_numbering looks like numbering
                        if self.looks_like_numbering(potential_numbering):
                            numbering = potential_numbering
                            content = potential_content
                            pattern_matched = "manual_detection"
                            level = self.determine_level_manual(potential_numbering)
                            break
            
            if numbering and content:
                text_lines.append(TextLineMatch(
                    line_number=i,
                    numbering=numbering,
                    content=content,
                    pattern_matched=pattern_matched,
                    level=level
                ))
        
        return text_lines
    
    def looks_like_numbering(self, text: str) -> bool:
        """Check if text looks like numbering"""
        # Remove common punctuation
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
    
    def determine_level_manual(self, numbering: str) -> Optional[int]:
        """Manually determine level from numbering text"""
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
    
    def match_numbering_to_content(self, text_lines: List[TextLineMatch], content_blocks: List[ContentBlock]) -> List[NumberingMatch]:
        """Match numbering from text to content blocks"""
        matches = []
        
        for text_line in text_lines:
            best_match = None
            best_confidence = 0.0
            
            for content_block in content_blocks:
                if content_block.block_type != "content":
                    continue
                
                # Try different matching strategies
                for strategy in self.matching_strategies:
                    confidence = self.calculate_match_confidence(
                        text_line, content_block, strategy
                    )
                    
                    if confidence > best_confidence:
                        best_confidence = confidence
                        best_match = NumberingMatch(
                            text_line=text_line,
                            content_block=content_block,
                            confidence=confidence,
                            match_type=strategy
                        )
            
            if best_match and best_match.confidence > 0.3:  # Minimum confidence threshold
                matches.append(best_match)
        
        return matches
    
    def calculate_match_confidence(self, text_line: TextLineMatch, content_block: ContentBlock, strategy: str) -> float:
        """Calculate confidence for a match using specified strategy"""
        
        if strategy == "exact_text_match":
            # Exact text match
            if text_line.content.lower() == content_block.text.lower():
                return 1.0
        
        elif strategy == "contains_text_match":
            # Content contains the text line content
            if text_line.content.lower() in content_block.text.lower():
                return 0.8
            # Text line content contains the content block text
            elif content_block.text.lower() in text_line.content.lower():
                return 0.7
        
        elif strategy == "fuzzy_text_match":
            # Fuzzy matching using common words
            text_words = set(text_line.content.lower().split())
            block_words = set(content_block.text.lower().split())
            
            if text_words and block_words:
                intersection = text_words.intersection(block_words)
                union = text_words.union(block_words)
                if union:
                    similarity = len(intersection) / len(union)
                    return similarity * 0.6
        
        elif strategy == "pattern_based_match":
            # Pattern-based matching (e.g., BWA- patterns)
            if "BWA-" in text_line.content and "BWA-" in content_block.text:
                return 0.4
        
        return 0.0
    
    def generate_matching_report(self, docx_path: str, txt_path: str) -> Dict[str, Any]:
        """Generate a comprehensive matching report"""
        
        # Extract content blocks from Word document
        print("Extracting content blocks from Word document...")
        content_blocks = self.extract_content_blocks_from_word(docx_path)
        print(f"Found {len(content_blocks)} content blocks")
        
        # Extract numbering from text file
        print("Extracting numbering from text file...")
        text_lines = self.extract_numbering_from_text(txt_path)
        print(f"Found {len(text_lines)} numbered lines")
        
        # Match numbering to content
        print("Matching numbering to content...")
        matches = self.match_numbering_to_content(text_lines, content_blocks)
        print(f"Found {len(matches)} matches")
        
        # Generate report
        report = {
            'content_blocks': [
                {
                    'index': block.index,
                    'text': block.text,
                    'block_type': block.block_type
                }
                for block in content_blocks
            ],
            'text_lines': [
                {
                    'line_number': line.line_number,
                    'numbering': line.numbering,
                    'content': line.content,
                    'pattern_matched': line.pattern_matched,
                    'level': line.level
                }
                for line in text_lines
            ],
            'matches': [
                {
                    'text_line': {
                        'line_number': match.text_line.line_number,
                        'numbering': match.text_line.numbering,
                        'content': match.text_line.content,
                        'level': match.text_line.level
                    },
                    'content_block': {
                        'index': match.content_block.index,
                        'text': match.content_block.text,
                        'block_type': match.content_block.block_type
                    },
                    'confidence': match.confidence,
                    'match_type': match.match_type
                }
                for match in matches
            ],
            'summary': {
                'total_content_blocks': len(content_blocks),
                'total_text_lines': len(text_lines),
                'total_matches': len(matches),
                'match_rate': len(matches) / len(text_lines) if text_lines else 0,
                'average_confidence': sum(match.confidence for match in matches) / len(matches) if matches else 0
            }
        }
        
        return report
    
    def print_matching_summary(self, report: Dict[str, Any]):
        """Print a summary of the matching results"""
        summary = report['summary']
        
        print(f"\n=== NUMBERING MATCHING SUMMARY ===")
        print(f"Content blocks from Word: {summary['total_content_blocks']}")
        print(f"Numbered lines from text: {summary['total_text_lines']}")
        print(f"Successful matches: {summary['total_matches']}")
        print(f"Match rate: {summary['match_rate']:.2%}")
        print(f"Average confidence: {summary['average_confidence']:.2f}")
        
        print(f"\n=== TOP MATCHES ===")
        matches = sorted(report['matches'], key=lambda x: x['confidence'], reverse=True)
        for i, match in enumerate(matches[:10]):
            print(f"{i+1:2d}. [{match['confidence']:.2f}] '{match['text_line']['numbering']}' → '{match['content_block']['text'][:50]}'")
        
        print(f"\n=== UNMATCHED TEXT LINES ===")
        matched_line_numbers = {match['text_line']['line_number'] for match in matches}
        unmatched_lines = [line for line in report['text_lines'] if line['line_number'] not in matched_line_numbers]
        for line in unmatched_lines[:10]:
            print(f"Line {line['line_number']}: '{line['numbering']}' → '{line['content']}'")
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save the matching report"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"Matching report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python numbering_pattern_matcher.py <docx_file> <txt_file> [output_dir]")
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
    
    matcher = NumberingPatternMatcher()
    
    try:
        print(f"Matching numbering patterns from: {txt_path}")
        print(f"To content blocks in: {docx_path}")
        
        # Generate matching report
        report = matcher.generate_matching_report(docx_path, txt_path)
        
        # Print summary
        matcher.print_matching_summary(report)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_numbering_matches.json")
        matcher.save_report(report, output_path)
        
        print(f"\nMatching report saved to: {output_path}")
        
    except Exception as e:
        print(f"Error matching numbering patterns: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 