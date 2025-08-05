# C# Open XML SDK Implementation

## Overview

This C# implementation uses the official Microsoft Open XML SDK to create Word documents with proper numbering. This approach is more robust and reliable than the Python XML manipulation approach.

## Prerequisites

- .NET 6.0 SDK or later
- Visual Studio 2022 or VS Code with C# extension

## Building the Project

### Using PowerShell:
```powershell
.\build_csharp.ps1
```

### Using Batch:
```cmd
build_csharp.bat
```

### Manual Build:
```bash
cd src
dotnet restore
dotnet build --configuration Release
dotnet publish --configuration Release --output ../output
```

## Usage

After building, the executable will be available at `output/WordNumberingRebuilder.exe`:

```bash
output/WordNumberingRebuilder.exe "input.json" "template.docx" "output.docx"
```

### Example:
```bash
output/WordNumberingRebuilder.exe "output/SECTION_00_00_00_hybrid_analysis.json" "output/complete_accuracy_check-fixed3.docx" "output/csharp_rebuilt.docx"
```

## Key Advantages

### 1. **Strongly-Typed Classes**
- Uses `Paragraph`, `NumberingProperties`, `AbstractNum`, etc.
- Compile-time safety - no typos in XML tag names
- IntelliSense support in Visual Studio

### 2. **Built-in Parts Management**
- No manual ZIP/unzip operations
- Automatic handling of `[Content_Types].xml` and relationships
- Proper part creation and linking

### 3. **Better Validation**
- SDK validates document structure
- Clear error messages for invalid operations
- Proper namespace handling

### 4. **Professional Tooling**
- Visual Studio integration
- Open XML SDK Productivity Tool for inspection
- Generate C# classes from existing documents

## Architecture

### Core Classes

#### `ParagraphInfo`
Represents a paragraph from the JSON analysis:
```csharp
public class ParagraphInfo
{
    public string Text { get; set; }
    public string? ListNumber { get; set; }
    public string? InferredNumber { get; set; }
    public int? Level { get; set; }
    public string? CleanedContent { get; set; }
    public bool IsListItem => !string.IsNullOrEmpty(ListNumber) || !string.IsNullOrEmpty(InferredNumber);
}
```

#### `ListDefinition`
Defines numbering format for each level:
```csharp
public class ListDefinition
{
    public int Level { get; set; }
    public string NumFmt { get; set; } = "decimal";
    public string LvlText { get; set; } = "%1.";
    public int Indent { get; set; } = 720;
    public int Hanging { get; set; } = 360;
}
```

#### `WordNumberingRebuilder`
Main class that handles document creation:
```csharp
public class WordNumberingRebuilder
{
    public void Rebuild(string jsonPath, string templatePath, string outputPath)
    {
        // 1. Load JSON data
        // 2. Copy template
        // 3. Create numbering definitions
        // 4. Build paragraphs with proper numbering
        // 5. Save document
    }
}
```

## How It Works

### 1. **JSON Parsing**
- Loads the JSON analysis file
- Maps JSON properties to `ParagraphInfo` objects
- Handles both true numbering and inferred numbering

### 2. **Template Copy**
- Uses a working Word document as template
- Preserves all styles, themes, and document structure
- Only replaces content and numbering

### 3. **Numbering Definitions**
- Creates separate `AbstractNum` for each level
- Defines proper formatting (decimal, lowerLetter, lowerRoman)
- Sets correct indentation and hanging values

### 4. **Paragraph Creation**
- Creates `Paragraph` objects with proper structure
- Adds `NumberingProperties` for list items
- Uses cleaned content when available

### 5. **Document Assembly**
- SDK handles all ZIP operations automatically
- Maintains proper relationships between parts
- Ensures valid Word document structure

## Numbering Levels

The implementation supports multiple numbering levels:

| Level | Format | Example | Indent |
|-------|--------|---------|--------|
| 0 | Decimal | 1. | 720 |
| 1 | Decimal | 1.01 | 1440 |
| 2 | Lower Letter | a. | 2160 |
| 3 | Decimal | 1. | 2880 |
| 4 | Lower Letter | a. | 3600 |
| 5 | Lower Roman | i. | 4320 |

## Error Handling

The C# implementation provides:
- Clear error messages for missing files
- Proper exception handling
- Validation of input parameters
- Detailed stack traces for debugging

## Integration with Python Pipeline

This C# implementation can be integrated into the existing Python pipeline:

1. **Python**: Generate JSON analysis
2. **C#**: Rebuild Word document with proper numbering
3. **Python**: Optional sanitization if needed

### Example Pipeline:
```bash
# 1. Python analysis
python src/enhanced_hybrid_detector.py "input.docx" "output.json"

# 2. C# rebuild
output/WordNumberingRebuilder.exe "output.json" "template.docx" "final_output.docx"

# 3. Optional Python sanitization
python src/docx_sanitizer.py "final_output.docx" "final_output_cleaned.docx"
```

## Benefits Over Python Approach

1. **Reliability**: SDK handles edge cases and validation
2. **Performance**: Native .NET performance
3. **Maintainability**: Strong typing and IntelliSense
4. **Professional**: Industry-standard approach
5. **Future-proof**: Microsoft-supported SDK

## Next Steps

1. Build the project using the provided scripts
2. Test with existing JSON files
3. Compare output quality with Python version
4. Integrate into production pipeline
5. Add additional features (headers, footers, tables, etc.)

The C# Open XML SDK approach provides a more robust and maintainable solution for Word document creation with proper numbering. 