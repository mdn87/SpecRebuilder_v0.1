using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        public void Rebuild(string jsonPath, string templatePath, string outputPath)
        {
            Console.WriteLine($"Loading JSON analysis from: {jsonPath}");
            
            // Load JSON data
            var jsonText = File.ReadAllText(jsonPath);
            var jsonDoc = JsonDocument.Parse(jsonText);
            var paragraphs = ParseJsonParagraphs(jsonDoc);
            
            Console.WriteLine($"Found {paragraphs.Count} paragraphs to process");
            
            // Copy template to output
            File.Copy(templatePath, outputPath, overwrite: true);
            
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null)
                {
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
            return element.TryGetProperty(propertyName, out var prop) ? prop.GetInt32() : null;
        }

        private Dictionary<int, uint> CreateNumberingDefinitions(NumberingDefinitionsPart numberingPart, List<ParagraphInfo> paragraphs)
        {
            var numberingMap = new Dictionary<int, uint>();
            uint abstractId = 1;
            uint numId = 1;

            // Get distinct levels that need numbering
            var levelsUsed = paragraphs
                .Where(p => p.IsListItem)
                .Select(p => p.Level ?? 0)
                .Distinct()
                .OrderBy(l => l)
                .ToList();

            foreach (var level in levelsUsed)
            {
                var listDef = GetListDefinition(level);
                
                // Create AbstractNum
                var abstractNum = new AbstractNum { AbstractNumberId = (UInt32Value)abstractId };
                var levelElement = new Level
                {
                    LevelIndex = (UInt32Value)(uint)level,
                    StartNumberingValue = new StartNumberingValue { Val = (UInt32Value)1 }
                };
                
                levelElement.AppendChild(new NumberingFormat { Val = GetNumberFormatValue(listDef.NumFmt) });
                levelElement.AppendChild(new LevelText { Val = listDef.LvlText });
                levelElement.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
                
                var pPr = new ParagraphProperties();
                var ind = new Indentation
                {
                    Left = listDef.Indent.ToString(),
                    Hanging = listDef.Hanging.ToString()
                };
                pPr.AppendChild(ind);
                levelElement.AppendChild(pPr);
                
                abstractNum.AppendChild(levelElement);
                numberingPart.Numbering.AppendChild(abstractNum);

                // Create NumberingInstance
                var num = new NumberingInstance { NumberID = (UInt32Value)numId };
                num.AppendChild(new AbstractNumId { Val = (UInt32Value)abstractId });
                numberingPart.Numbering.AppendChild(num);

                numberingMap[level] = numId;
                abstractId++;
                numId++;
            }

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
                0 => new ListDefinition { Level = 0, NumFmt = "decimal", LvlText = "%1.", Indent = 720, Hanging = 360 },
                1 => new ListDefinition { Level = 1, NumFmt = "decimal", LvlText = "%1.%2.", Indent = 1440, Hanging = 360 },
                2 => new ListDefinition { Level = 2, NumFmt = "lowerLetter", LvlText = "%1.", Indent = 2160, Hanging = 360 },
                3 => new ListDefinition { Level = 3, NumFmt = "decimal", LvlText = "%1.", Indent = 2880, Hanging = 360 },
                4 => new ListDefinition { Level = 4, NumFmt = "lowerLetter", LvlText = "%1.", Indent = 3600, Hanging = 360 },
                5 => new ListDefinition { Level = 5, NumFmt = "lowerRoman", LvlText = "%1.", Indent = 4320, Hanging = 360 },
                _ => new ListDefinition { Level = level, NumFmt = "decimal", LvlText = "%1.", Indent = 720 * (level + 1), Hanging = 360 }
            };
        }

        private Paragraph CreateParagraph(ParagraphInfo paraInfo, Dictionary<int, uint> numberingMap)
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties();

            // Add numbering properties if this is a list item
            if (paraInfo.IsListItem && paraInfo.Level.HasValue)
            {
                var level = paraInfo.Level.Value;
                if (numberingMap.TryGetValue(level, out var numId))
                {
                    var numPr = new NumberingProperties(
                        new NumberingLevelReference { Val = (UInt32Value)(uint)level },
                        new NumberingId { Val = (UInt32Value)numId }
                    );
                    pPr.AppendChild(numPr);
                }
            }

            paragraph.AppendChild(pPr);

            // Add the text run
            var run = new Run();
            var text = new Text(paraInfo.DisplayText) { Space = SpaceProcessingModeValues.Preserve };
            run.AppendChild(text);
            paragraph.AppendChild(run);

            return paragraph;
        }
    }

    /// <summary>
    /// Main program entry point
    /// </summary>
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: WordNumberingRebuilder.exe <json_file> <template_docx> <output_docx>");
                Console.WriteLine("Example: WordNumberingRebuilder.exe output/SECTION_00_00_00_hybrid_analysis.json output/complete_accuracy_check-fixed3.docx output/word_numbering_rebuilt.docx");
                return;
            }

            var jsonPath = args[0];
            var templatePath = args[1];
            var outputPath = args[2];

            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON file not found: {jsonPath}");
                return;
            }

            if (!File.Exists(templatePath))
            {
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
                Console.WriteLine($"Error rebuilding document: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
} 