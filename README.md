# SpecRebuilder v0.1 - Multilist Level Formatting Analyzer

A fresh approach to detecting and analyzing multilist level formatting in Word documents. This tool converts Word documents to JSON format and analyzes the structure to identify formatting patterns, issues, and recommendations.

## Features

- **Word to JSON Conversion**: Extracts complete document structure including paragraphs, styles, numbering, headers, footers, and metadata
- **Multilist Analysis**: Analyzes numbering patterns, style consistency, and structural issues
- **Comprehensive Reporting**: Generates detailed analysis reports with warnings and recommendations
- **Pipeline Integration**: Combines conversion and analysis in a single command

## Quick Start

### Analyze a Word Document

```bash
python src/analyze_document.py "path/to/document.docx"
```

This will:
1. Convert the Word document to JSON format
2. Analyze the multilist structure
3. Generate a comprehensive analysis report
4. Display summary, warnings, and recommendations

### Individual Tools

#### Convert Word to JSON
```bash
python src/word_to_json.py "path/to/document.docx" [output.json]
```

#### Analyze JSON Structure
```bash
python src/multilist_analyzer.py "document_structure.json" [analysis.json]
```

## Analysis Results

The tool provides detailed analysis including:

- **Summary Statistics**: Total levels, numbered vs unnumbered content, unique styles and numbering IDs
- **Numbering Analysis**: Breakdown of numbering schemes and their usage
- **Style Analysis**: How different styles are used and their numbering associations
- **Warnings**: Potential formatting issues and inconsistencies
- **Recommendations**: Suggestions for improving document structure

## Example Output

```
=== ANALYSIS SUMMARY ===
Document: examples/SECTION 00 00 00.docx
Total levels: 41
Numbered levels: 16
Unnumbered levels: 25
Unique numbering IDs: 1
Unique styles: 5
Errors: 0
Warnings: 1

=== WARNINGS ===
  ⚠️  Found 25 levels without numbering

=== RECOMMENDATIONS ===
  💡 Review warnings for potential formatting issues
  💡 High percentage of unnumbered content - consider adding numbering

=== NUMBERING ANALYSIS ===
  numbering_id_10: 16 levels, styles: ['LEVEL 4 - JE']

=== STYLE ANALYSIS ===
  Normal: 3 uses (no numbering)
  LEVEL 1 - JE: 2 uses (no numbering)
  LEVEL 2 - JE: 4 uses (no numbering)
  LEVEL 3 - JE: 8 uses (no numbering)
```

## What We've Learned

From analyzing the test document `SECTION 00 00 00.docx`:

1. **Structure Detection**: Successfully extracted 41 content levels with 5 different styles
2. **Numbering Patterns**: Found 1 numbering scheme (ID: 10) with 16 numbered levels
3. **Style Analysis**: Identified clear style hierarchy (Normal → LEVEL 1-4)
4. **Issues Detected**: 25 levels lack proper numbering, indicating potential formatting problems
5. **Recommendations**: Document would benefit from consistent numbering application

## Key Insights

- **Style-Based Hierarchy**: The document uses a clear style hierarchy (LEVEL 1-4) for different content levels
- **Mixed Numbering**: Only some content has numbering applied, suggesting inconsistent formatting
- **Single Numbering Scheme**: All numbered content uses the same numbering ID (10), indicating a single list structure
- **Level Gaps**: Numbering levels jump from 4 to 5, suggesting some levels may be missing

## Next Steps

This analysis provides a foundation for:
1. **Formatting Validation**: Identifying documents with broken or inconsistent formatting
2. **Structure Repair**: Understanding what needs to be fixed in problematic documents
3. **Template Analysis**: Comparing documents against expected formatting patterns
4. **Automated Fixes**: Using the analysis to guide automated formatting corrections

## Dependencies

- `python-docx`: Word document processing
- `lxml`: XML parsing for document structure
- Standard Python libraries: `json`, `pathlib`, `typing`, `dataclasses`

## File Structure

```
SpecRebuilder_v0.1/
├── src/
│   ├── word_to_json.py          # Word to JSON converter
│   ├── multilist_analyzer.py    # Structure analyzer
│   └── analyze_document.py      # Complete pipeline
├── examples/
│   └── SECTION 00 00 00.docx   # Test document
└── README.md                    # This file
```

## Usage Examples

### Basic Analysis
```bash
cd SpecRebuilder_v0.1
python src/analyze_document.py "examples/SECTION 00 00 00.docx"
```

### Custom Output Directory
```bash
python src/analyze_document.py "document.docx" "output/"
```

### Step-by-Step Analysis
```bash
# Step 1: Convert to JSON
python src/word_to_json.py "document.docx"

# Step 2: Analyze the JSON
python src/multilist_analyzer.py "document_structure.json"
```

This tool provides a solid foundation for understanding and fixing multilist formatting issues in Word documents. 