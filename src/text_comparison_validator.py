#!/usr/bin/env python3
"""
Text Comparison Validator

This script extracts text from Word documents and compares it with text files
to validate that our extraction is accurate for production use.
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
class TextComparison:
    """Represents a comparison between Word and text file content"""
    word_text: str
    text_file_content: str
    line_number: int
    is_exact_match: bool
    differences: List[str] = None

@dataclass
class ValidationResult:
    """Represents the overall validation result"""
    total_lines: int
    exact_matches: int
    partial_matches: int
    mismatches: int
    match_percentage: float
    details: List[TextComparison]

class TextComparisonValidator:
    """Validates text extraction from Word documents against text files"""
    
    def __init__(self):
        # Text cleaning strategies for comparison
        self.cleaning_strategies = [
            "normalize_whitespace",
            "remove_extra_spaces",
            "normalize_newlines",
            "strip_punctuation",
            "lowercase_comparison"
        ]
        
        # Comparison strategies
        self.comparison_strategies = [
            "exact_match",
            "normalized_match",
            "content_only_match",
            "fuzzy_match"
        ]
    
    def extract_text_from_word(self, docx_path: str) -> List[str]:
        """Extract all text from Word document"""
        # Convert Word to JSON
        converter = WordToJsonConverter()
        json_path = converter.convert_to_json(docx_path)
        
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        text_lines = []
        
        for paragraph in data.get('paragraphs', []):
            # Get paragraph text
            paragraph_text = paragraph.get('text', '').strip()
            
            # If paragraph text is empty, try to get from runs
            if not paragraph_text:
                runs_text = ""
                for run in paragraph.get('runs', []):
                    runs_text += run.get('text', '')
                paragraph_text = runs_text.strip()
            
            if paragraph_text:  # Only include non-empty paragraphs
                text_lines.append(paragraph_text)
        
        return text_lines
    
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
    
    def clean_text_for_comparison(self, text: str, strategy: str = "normalized_match") -> str:
        """Clean text for comparison based on strategy"""
        if strategy == "exact_match":
            return text
        
        elif strategy == "normalized_match":
            # Normalize whitespace and newlines
            text = re.sub(r'\s+', ' ', text)
            text = text.replace('\n', ' ').replace('\r', ' ')
            return text.strip()
        
        elif strategy == "content_only_match":
            # Remove punctuation and normalize
            text = re.sub(r'[^\w\s]', '', text)
            text = re.sub(r'\s+', ' ', text)
            return text.strip().lower()
        
        elif strategy == "fuzzy_match":
            # Most aggressive cleaning
            text = re.sub(r'[^\w\s]', '', text)
            text = re.sub(r'\s+', ' ', text)
            return text.strip().lower()
        
        return text
    
    def compare_texts(self, word_lines: List[str], text_lines: List[str]) -> ValidationResult:
        """Compare Word text with text file content"""
        comparisons = []
        exact_matches = 0
        partial_matches = 0
        mismatches = 0
        
        # Use the shorter list length to avoid index errors
        max_lines = min(len(word_lines), len(text_lines))
        
        for i in range(max_lines):
            word_text = word_lines[i]
            text_file_content = text_lines[i]
            
            # Try different comparison strategies
            is_exact_match = False
            differences = []
            
            # Strategy 1: Exact match
            if word_text == text_file_content:
                is_exact_match = True
                exact_matches += 1
            else:
                # Strategy 2: Normalized match
                normalized_word = self.clean_text_for_comparison(word_text, "normalized_match")
                normalized_text = self.clean_text_for_comparison(text_file_content, "normalized_match")
                
                if normalized_word == normalized_text:
                    partial_matches += 1
                    differences.append("Normalized match (whitespace differences)")
                else:
                    # Strategy 3: Content only match
                    content_word = self.clean_text_for_comparison(word_text, "content_only_match")
                    content_text = self.clean_text_for_comparison(text_file_content, "content_only_match")
                    
                    if content_word == content_text:
                        partial_matches += 1
                        differences.append("Content match (punctuation differences)")
                    else:
                        mismatches += 1
                        differences.append(f"Word: '{word_text[:50]}...' vs Text: '{text_file_content[:50]}...'")
            
            comparison = TextComparison(
                word_text=word_text,
                text_file_content=text_file_content,
                line_number=i + 1,
                is_exact_match=is_exact_match,
                differences=differences
            )
            comparisons.append(comparison)
        
        # Handle any remaining lines
        if len(word_lines) > len(text_lines):
            for i in range(len(text_lines), len(word_lines)):
                comparison = TextComparison(
                    word_text=word_lines[i],
                    text_file_content="",
                    line_number=i + 1,
                    is_exact_match=False,
                    differences=["Extra line in Word document"]
                )
                comparisons.append(comparison)
                mismatches += 1
        
        elif len(text_lines) > len(word_lines):
            for i in range(len(word_lines), len(text_lines)):
                comparison = TextComparison(
                    word_text="",
                    text_file_content=text_lines[i],
                    line_number=i + 1,
                    is_exact_match=False,
                    differences=["Extra line in text file"]
                )
                comparisons.append(comparison)
                mismatches += 1
        
        total_lines = len(comparisons)
        match_percentage = (exact_matches + partial_matches) / total_lines if total_lines > 0 else 0
        
        return ValidationResult(
            total_lines=total_lines,
            exact_matches=exact_matches,
            partial_matches=partial_matches,
            mismatches=mismatches,
            match_percentage=match_percentage,
            details=comparisons
        )
    
    def generate_validation_report(self, docx_path: str, txt_path: str) -> Dict[str, Any]:
        """Generate a comprehensive validation report"""
        
        # Extract text from Word document
        print("Extracting text from Word document...")
        word_lines = self.extract_text_from_word(docx_path)
        print(f"Found {len(word_lines)} text lines in Word document")
        
        # Read text file
        print("Reading text file...")
        text_lines = self.read_text_file(txt_path)
        print(f"Found {len(text_lines)} text lines in text file")
        
        # Compare texts
        print("Comparing texts...")
        result = self.compare_texts(word_lines, text_lines)
        print(f"Comparison complete: {result.exact_matches} exact matches, {result.partial_matches} partial matches, {result.mismatches} mismatches")
        
        # Generate report
        report = {
            'word_document': {
                'path': docx_path,
                'total_lines': len(word_lines),
                'sample_lines': word_lines[:10] if word_lines else []
            },
            'text_file': {
                'path': txt_path,
                'total_lines': len(text_lines),
                'sample_lines': text_lines[:10] if text_lines else []
            },
            'validation_result': {
                'total_lines': result.total_lines,
                'exact_matches': result.exact_matches,
                'partial_matches': result.partial_matches,
                'mismatches': result.mismatches,
                'match_percentage': result.match_percentage
            },
            'detailed_comparisons': [
                {
                    'line_number': comp.line_number,
                    'word_text': comp.word_text,
                    'text_file_content': comp.text_file_content,
                    'is_exact_match': comp.is_exact_match,
                    'differences': comp.differences
                }
                for comp in result.details
            ]
        }
        
        return report
    
    def print_validation_summary(self, report: Dict[str, Any]):
        """Print a summary of the validation results"""
        word_doc = report['word_document']
        text_file = report['text_file']
        result = report['validation_result']
        
        print(f"\n=== TEXT VALIDATION SUMMARY ===")
        print(f"Word document: {word_doc['path']}")
        print(f"  - Total lines: {word_doc['total_lines']}")
        print(f"Text file: {text_file['path']}")
        print(f"  - Total lines: {text_file['total_lines']}")
        
        print(f"\n=== COMPARISON RESULTS ===")
        print(f"Total lines compared: {result['total_lines']}")
        print(f"Exact matches: {result['exact_matches']}")
        print(f"Partial matches: {result['partial_matches']}")
        print(f"Mismatches: {result['mismatches']}")
        print(f"Match percentage: {result['match_percentage']:.2%}")
        
        print(f"\n=== SAMPLE COMPARISONS ===")
        comparisons = report['detailed_comparisons']
        for i, comp in enumerate(comparisons[:10]):
            status = "✓" if comp['is_exact_match'] else "△" if "match" in str(comp['differences']) else "✗"
            print(f"{i+1:2d}. [{status}] Line {comp['line_number']}: '{comp['word_text'][:50]}...'")
        
        print(f"\n=== MISMATCHES ===")
        mismatches = [comp for comp in comparisons if not comp['is_exact_match'] and "match" not in str(comp['differences'])]
        for i, comp in enumerate(mismatches[:10]):
            print(f"{i+1:2d}. Line {comp['line_number']}: {comp['differences']}")
        
        if result['match_percentage'] >= 0.95:
            print(f"\n✅ EXCELLENT: {result['match_percentage']:.1%} match rate - Text extraction is highly accurate!")
        elif result['match_percentage'] >= 0.80:
            print(f"\n⚠️  GOOD: {result['match_percentage']:.1%} match rate - Text extraction is mostly accurate with minor differences.")
        else:
            print(f"\n❌ POOR: {result['match_percentage']:.1%} match rate - Text extraction needs improvement.")
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save the validation report"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False, default=str)
        print(f"Validation report saved to: {output_path}")

def main():
    """Main function"""
    if len(sys.argv) < 3:
        print("Usage: python text_comparison_validator.py <docx_file> <txt_file> [output_dir]")
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
    
    validator = TextComparisonValidator()
    
    try:
        print(f"Validating text extraction from: {docx_path}")
        print(f"Comparing with text file: {txt_path}")
        
        # Generate validation report
        report = validator.generate_validation_report(docx_path, txt_path)
        
        # Print summary
        validator.print_validation_summary(report)
        
        # Save report
        base_name = Path(docx_path).stem
        output_path = os.path.join(output_dir, f"{base_name}_text_validation.json")
        validator.save_report(report, output_path)
        
        print(f"\nValidation report saved to: {output_path}")
        
    except Exception as e:
        print(f"Error in text validation: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 