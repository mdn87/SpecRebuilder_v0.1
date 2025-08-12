using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SpecRebuilder
{
    /// <summary>
    /// Represents a paragraph from the JSON analysis
    /// </summary>
    public class ParagraphInfo
    {
        public string Text { get; set; } = "";
        public string? ListNumber { get; set; }
        public string? InferredNumber { get; set; }
        public int? Level { get; set; }
        public string? CleanedContent { get; set; }
        public string? NumberingType { get; set; }
        
        public bool IsListItem => !string.IsNullOrEmpty(ListNumber) || !string.IsNullOrEmpty(InferredNumber);
        public string DisplayText => !string.IsNullOrEmpty(CleanedContent) ? CleanedContent : Text;
    }

    /// <summary>
    /// Represents a list definition for numbering
    /// </summary>
    public class ListDefinition
    {
        public int Level { get; set; }
        public string NumFmt { get; set; } = "decimal";
        public string LvlText { get; set; } = "%1.";
        public int Indent { get; set; } = 720;
        public int Hanging { get; set; } = 360;
    }

    /// <summary>
    /// Word Numbering Rebuilder using Open XML SDK
    /// </summary>
    public class WordNumberingRebuilder
    {
        public static StreamWriter? _logWriter;
        
        public static void Log(string message)
        {
            string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            Console.WriteLine(logLine);
            _logWriter?.WriteLine(logLine);
            _logWriter?.Flush();
        }
        
        public void Rebuild(string jsonPath, string templatePath, string outputPath)
        {
            string logPath = Path.Combine(Path.GetDirectoryName(outputPath) ?? ".", "WordNumberingRebuilder.txt");
            using (_logWriter = new StreamWriter(logPath, append: true))
            {
                Console.WriteLine($"Loading JSON analysis from: {jsonPath}");
                Log($"Loading JSON analysis from: {jsonPath}");

                // Load JSON data
                var jsonText = File.ReadAllText(jsonPath);
                var jsonDoc = JsonDocument.Parse(jsonText);
                var paragraphs = ParseJsonParagraphs(jsonDoc);
                
                Console.WriteLine($"Found {paragraphs.Count} paragraphs to process");
                Log($"Found {paragraphs.Count} paragraphs to process");

                // Copy template to output
                File.Copy(templatePath, outputPath, overwrite: true);
                
                using (var doc = WordprocessingDocument.Open(outputPath, true))
                {
                    var mainPart = doc.MainDocumentPart;
                    if (mainPart == null)
                    {
                        Log("Main document part is null");
                        throw new InvalidOperationException("Main document part is null");
                    }
                    
                    var numberingPart = mainPart.NumberingDefinitionsPart 
                                      ?? mainPart.AddNewPart<NumberingDefinitionsPart>();

                    // Clear existing numbering
                    numberingPart.Numbering = new Numbering();

                    // Create numbering definitions
                    var numberingMap = CreateNumberingDefinitions(numberingPart, paragraphs);
                    
                    // Clear and rebuild document body
                    var body = mainPart.Document.Body;
                    if (body == null)
                    {
                        Log("Document body is null");
                        throw new InvalidOperationException("Document body is null");
                    }
                    
                    body.RemoveAllChildren();
                    
                    foreach (var paraInfo in paragraphs)
                    {
                        var paragraph = CreateParagraph(paraInfo, numberingMap);
                        body.AppendChild(paragraph);
                    }

                    mainPart.Document.Save();
                }
                
                Console.WriteLine($"Document saved to: {outputPath}");
                Console.WriteLine("Document rebuild complete!");
                Log($"Document saved to: {outputPath}");
                Log("Document rebuild complete!");
            }
        }

        private List<ParagraphInfo> ParseJsonParagraphs(JsonDocument jsonDoc)
        {
            var paragraphs = new List<ParagraphInfo>();
            
            if (jsonDoc.RootElement.TryGetProperty("all_paragraphs", out var parasElement))
            {
                foreach (var paraElement in parasElement.EnumerateArray())
                {
                    var paraInfo = new ParagraphInfo
                    {
                        Text = GetStringProperty(paraElement, "text"),
                        ListNumber = GetStringProperty(paraElement, "list_number"),
                        InferredNumber = GetStringProperty(paraElement, "inferred_number"),
                        Level = GetIntProperty(paraElement, "level"),
                        CleanedContent = GetStringProperty(paraElement, "cleaned_content"),
                        NumberingType = GetStringProperty(paraElement, "numbering_type")
                    };
                    
                    // If no cleaned_content, use the text field
                    if (string.IsNullOrEmpty(paraInfo.CleanedContent))
                    {
                        paraInfo.CleanedContent = paraInfo.Text;
                    }
                    
                    paragraphs.Add(paraInfo);
                }
            }
            
            return paragraphs;
        }

        private string GetStringProperty(JsonElement element, string propertyName)
        {
            return element.TryGetProperty(propertyName, out var prop) ? prop.GetString() ?? "" : "";
        }

        private int? GetIntProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                if (prop.ValueKind == JsonValueKind.Null)
                    return null;
                if (prop.ValueKind == JsonValueKind.Number)
                    return prop.GetInt32();
            }
            return null;
        }

        private Dictionary<int, uint> CreateNumberingDefinitions(NumberingDefinitionsPart numberingPart, List<ParagraphInfo> paragraphs)
        {
            var numberingMap = new Dictionary<int, uint>();
            var nextNumId = 1U;
            
            // Get distinct levels that need numbering
            var levelsUsed = paragraphs
                .Where(p => p.IsListItem)
                .Select(p => p.Level ?? 0)
                .Distinct()
                .OrderBy(l => l)
                .ToList();

            Log($"Creating numbering for levels: {string.Join(", ", levelsUsed)}");
            Console.WriteLine($"Creating numbering for levels: {string.Join(", ", levelsUsed)}");

            // Create separate numbering instances for each level to allow proper sequencing
            foreach (var level in levelsUsed)
            {
                var listDef = GetListDefinition(level);
                
                // Create AbstractNum for this level
                var abstractId = nextNumId;
                var abstractNum = new AbstractNum { AbstractNumberId = (Int32Value)(int)abstractId };
                
                // Add required properties for multi-level lists
                abstractNum.AppendChild(new Nsid { Val = (HexBinaryValue)("12345678" + level.ToString("X4")) }); // Unique ID per level
                abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
                abstractNum.AppendChild(new TemplateCode { Val = (HexBinaryValue)("0409" + level.ToString("X4")) });
                
                // Create level definition
                var levelElement = new Level
                {
                    LevelIndex = (Int32Value)0, // Always 0 for single-level lists
                    TemplateCode = (HexBinaryValue)("0409" + level.ToString("X4")),
                    StartNumberingValue = new StartNumberingValue { Val = (Int32Value)1 }
                };
                
                levelElement.AppendChild(new NumberingFormat { Val = GetNumberFormatValue(listDef.NumFmt) });
                levelElement.AppendChild(new LevelText { Val = "%1." }); // Simple %1. for single-level
                levelElement.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
                
                // Add indentation based on level
                var ind = new Indentation
                {
                    Left = (StringValue)(level * 720).ToString(), // 720 twips = 0.5 inch per level
                    Hanging = (StringValue)"360"
                };
                levelElement.AppendChild(ind);
                
                abstractNum.AppendChild(levelElement);
                numberingPart.Numbering.AppendChild(abstractNum);

                // Create NumberingInstance for this level
                var num = new NumberingInstance { NumberID = (Int32Value)(int)nextNumId };
                num.AppendChild(new AbstractNumId { Val = (Int32Value)(int)abstractId });
                numberingPart.Numbering.AppendChild(num);

                // Map this level to its numbering instance
                numberingMap[level] = nextNumId;
                
                Log($"  Level {level}: Format={listDef.NumFmt}, Text=%1., numId={nextNumId}");
                Console.WriteLine($"  Level {level}: Format={listDef.NumFmt}, Text=%1., numId={nextNumId}");
                
                nextNumId++;
            }

            Log($"Created {levelsUsed.Count} separate numbering instances for proper sequencing");
            Console.WriteLine($"Created {levelsUsed.Count} separate numbering instances for proper sequencing");

            return numberingMap;
        }

        private NumberFormatValues GetNumberFormatValue(string numFmt)
        {
            return numFmt switch
            {
                "decimal" => NumberFormatValues.Decimal,
                "lowerLetter" => NumberFormatValues.LowerLetter,
                "upperLetter" => NumberFormatValues.UpperLetter,
                "lowerRoman" => NumberFormatValues.LowerRoman,
                "upperRoman" => NumberFormatValues.UpperRoman,
                _ => NumberFormatValues.Decimal
            };
        }

        private ListDefinition GetListDefinition(int level)
        {
            return level switch
            {
                1 => new ListDefinition { Level = 1, NumFmt = "decimal", LvlText = "%1.", Indent = 720, Hanging = 360 },
                2 => new ListDefinition { Level = 2, NumFmt = "decimal", LvlText = "%1.%2.", Indent = 1440, Hanging = 360 },
                3 => new ListDefinition { Level = 3, NumFmt = "upperLetter", LvlText = "%1.", Indent = 2160, Hanging = 360 },
                4 => new ListDefinition { Level = 4, NumFmt = "decimal", LvlText = "%1.", Indent = 2880, Hanging = 360 },
                5 => new ListDefinition { Level = 5, NumFmt = "lowerLetter", LvlText = "%1.", Indent = 3600, Hanging = 360 },
                6 => new ListDefinition { Level = 6, NumFmt = "lowerRoman", LvlText = "%1.", Indent = 4320, Hanging = 360 },
                _ => new ListDefinition { Level = level, NumFmt = "decimal", LvlText = "%1.", Indent = 720 * level, Hanging = 360 }
            };
        }

        private Paragraph CreateParagraph(ParagraphInfo paraInfo, Dictionary<int, uint> numberingMap)
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties();

            // Add numbering properties directly if this is a list item
            if (paraInfo.IsListItem && paraInfo.Level.HasValue)
            {
                var level = paraInfo.Level.Value;
                if (numberingMap.TryGetValue(level, out var numId))
                {
                    var numPr = new NumberingProperties();
                    numPr.AppendChild(new NumberingLevelReference { Val = (Int32Value)0 }); // Always 0 for single-level
                    numPr.AppendChild(new NumberingId { Val = (Int32Value)(int)numId });
                    
                    // Add start value if we have a specific number from the JSON
                    if (!string.IsNullOrEmpty(paraInfo.ListNumber))
                    {
                        // Try to extract the numeric value from the list number
                        var startValue = ExtractStartValue(paraInfo.ListNumber);
                        if (startValue.HasValue)
                        {
                            numPr.AppendChild(new NumberingLevelReference { Val = (Int32Value)startValue.Value });
                        }
                    }
                    
                    pPr.AppendChild(numPr);
                    Log($"Applied numbering level {level} with numId {numId} to: {paraInfo.DisplayText} (ListNumber: {paraInfo.ListNumber})");
                    Console.WriteLine($"Applied numbering level {level} with numId {numId} to: {paraInfo.DisplayText} (ListNumber: {paraInfo.ListNumber})");
                }
                else
                {
                    Log($"No numbering found for level {level}: {paraInfo.DisplayText}");
                    Console.WriteLine($"No numbering found for level {level}: {paraInfo.DisplayText}");
                }
            }
            else
            {
                Log($"Not a list item: {paraInfo.DisplayText}");
                Console.WriteLine($"Not a list item: {paraInfo.DisplayText}");
            }

            paragraph.AppendChild(pPr);

            // Add the text run
            var run = new Run();
            var text = new Text(paraInfo.DisplayText) { Space = SpaceProcessingModeValues.Preserve };
            run.AppendChild(text);
            paragraph.AppendChild(run);

            return paragraph;
        }
        
        private int? ExtractStartValue(string listNumber)
        {
            if (string.IsNullOrEmpty(listNumber))
                return null;
                
            // Extract numeric value from patterns like "1.0", "1.01", "A.", "1.", etc.
            var match = System.Text.RegularExpressions.Regex.Match(listNumber, @"^(\d+)");
            if (match.Success && int.TryParse(match.Groups[1].Value, out var value))
            {
                return value;
            }
            
            // For letter patterns, convert to number (A=1, B=2, etc.)
            var letterMatch = System.Text.RegularExpressions.Regex.Match(listNumber, @"^([A-Z])");
            if (letterMatch.Success)
            {
                var letter = letterMatch.Groups[1].Value[0];
                return letter - 'A' + 1;
            }
            
            // For lowercase letters
            var lowerLetterMatch = System.Text.RegularExpressions.Regex.Match(listNumber, @"^([a-z])");
            if (lowerLetterMatch.Success)
            {
                var letter = lowerLetterMatch.Groups[1].Value[0];
                return letter - 'a' + 1;
            }
            
            return null;
        }
    }

    /// <summary>
    /// Main program entry point
    /// </summary>
    public class Program
    {
        public static StreamWriter? _logWriter;
        
        public static void Log(string message)
        {
            string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            Console.WriteLine(logLine);
            _logWriter?.WriteLine(logLine);
            _logWriter?.Flush();
        }
        
        public static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Log("Usage: WordNumberingRebuilder.exe <json_file> <template_docx> <output_docx>");
                Console.WriteLine("Usage: WordNumberingRebuilder.exe <json_file> <template_docx> <output_docx>");
                Log("Example: WordNumberingRebuilder.exe output/SECTION_00_00_00_hybrid_analysis.json output/complete_accuracy_check-fixed3.docx output/word_numbering_rebuilt.docx");
                Console.WriteLine("Example: WordNumberingRebuilder.exe output/SECTION_00_00_00_hybrid_analysis.json output/complete_accuracy_check-fixed3.docx output/word_numbering_rebuilt.docx");
                return;
            }

            var jsonPath = args[0];
            var templatePath = args[1];
            var outputPath = args[2];

            if (!File.Exists(jsonPath))
            {
                Log($"Error: JSON file not found: {jsonPath}");
                Console.WriteLine($"Error: JSON file not found: {jsonPath}");
                return;
            }

            if (!File.Exists(templatePath))
            {
                Log($"Error: Template file not found: {templatePath}");
                Console.WriteLine($"Error: Template file not found: {templatePath}");
                return;
            }

            // Create output directory if it doesn't exist
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                var rebuilder = new WordNumberingRebuilder();
                rebuilder.Rebuild(jsonPath, templatePath, outputPath);
            }
            catch (Exception ex)
            {
                Log($"Error rebuilding document: {ex.Message}");
                Console.WriteLine($"Error rebuilding document: {ex.Message}");
                //Log(ex.StackTrace);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}