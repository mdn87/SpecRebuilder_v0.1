using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordNumberingRebuilder
{
    public class EnhancedBlock
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public string CleanedContent { get; set; }
        public int? Level { get; set; }
        public string NumFmt { get; set; }
        public int? ListId { get; set; }
        public int? ParentListId { get; set; }
        public string NumberingPattern { get; set; }
        public string InferredNumber { get; set; }
        public bool IsListItem { get; set; }
        public bool IsContinuation { get; set; }
        public int? ContinuationOf { get; set; }
        public int? IndentationLevel { get; set; }
        public List<string> ContextHints { get; set; }
        public double ConfidenceScore { get; set; }
    }

    public class EnhancedAnalysis
    {
        public Dictionary<string, object> Analysis { get; set; }
        public List<EnhancedBlock> EnhancedBlocks { get; set; }
        public List<List<int>> ListGroups { get; set; }
    }

    public class EnhancedWordRebuilder
    {
        public void RebuildDocument(string jsonPath, string templatePath, string outputPath)
        {
            Console.WriteLine($"Loading enhanced analysis from: {jsonPath}");
            
            // Load enhanced analysis
            var analysis = LoadEnhancedAnalysis(jsonPath);
            var blocks = analysis.EnhancedBlocks;
            
            Console.WriteLine($"Found {blocks.Count} enhanced blocks to process");
            
            // Create document from template
            File.Copy(templatePath, outputPath, true);
            
            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                var mainPart = doc.MainDocumentPart;
                var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
                
                // Clear existing numbering
                numberingPart.Numbering = new Numbering();
                
                // Create numbering definitions for each list group
                var numberingMap = CreateEnhancedNumberingDefinitions(numberingPart, blocks, analysis.ListGroups);
                
                // Clear and rebuild document body
                var body = mainPart.Document.Body;
                body.RemoveAllChildren();
                
                // Create paragraphs from enhanced blocks
                foreach (var block in blocks)
                {
                    var paragraph = CreateEnhancedParagraph(block, numberingMap);
                    body.AppendChild(paragraph);
                }
                
                doc.Save();
            }
            
            Console.WriteLine($"Enhanced document saved to: {outputPath}");
            Console.WriteLine("Enhanced document rebuild complete!");
        }
        
        private EnhancedAnalysis LoadEnhancedAnalysis(string jsonPath)
        {
            var json = File.ReadAllText(jsonPath);
            var analysis = JsonSerializer.Deserialize<EnhancedAnalysis>(json);
            return analysis;
        }
        
        private Dictionary<int, uint> CreateEnhancedNumberingDefinitions(
            NumberingDefinitionsPart numberingPart, 
            List<EnhancedBlock> blocks, 
            List<List<int>> listGroups)
        {
            var numberingMap = new Dictionary<int, uint>();
            var nextNumId = 1U;
            
            // Group blocks by list ID
            var blocksByList = new Dictionary<int, List<EnhancedBlock>>();
            foreach (var block in blocks)
            {
                if (block.ListId.HasValue)
                {
                    if (!blocksByList.ContainsKey(block.ListId.Value))
                        blocksByList[block.ListId.Value] = new List<EnhancedBlock>();
                    blocksByList[block.ListId.Value].Add(block);
                }
            }
            
            // Create numbering definition for each list group
            foreach (var kvp in blocksByList)
            {
                var listId = kvp.Key;
                var listBlocks = kvp.Value;
                
                Console.WriteLine($"Creating numbering for list {listId} with {listBlocks.Count} blocks");
                
                // Get unique levels in this list
                var levels = new HashSet<int>();
                foreach (var block in listBlocks)
                {
                    if (block.Level.HasValue)
                        levels.Add(block.Level.Value);
                }
                
                // Create AbstractNum for this list
                var abstractNum = CreateAbstractNum(nextNumId, levels, listBlocks);
                numberingPart.Numbering.AppendChild(abstractNum);
                
                // Create NumberingInstance
                var numberingInstance = new NumberingInstance
                {
                    NumberID = (UInt32Value)nextNumId,
                    AbstractNumId = new AbstractNumId { Val = (Int32Value)(int)nextNumId }
                };
                numberingPart.Numbering.AppendChild(numberingInstance);
                
                // Map levels to this numbering instance
                foreach (var level in levels)
                {
                    numberingMap[level] = (uint)nextNumId;
                }
                
                nextNumId++;
            }
            
            return numberingMap;
        }
        
        private AbstractNum CreateAbstractNum(uint numId, HashSet<int> levels, List<EnhancedBlock> listBlocks)
        {
            var abstractNum = new AbstractNum
            {
                AbstractNumberId = (UInt32Value)numId,
                Nsid = new Nsid { Val = "12345678" },
                MultiLevelType = (MultiLevelType)MultiLevelValues.HybridMultilevel
            };
            
            // Create level definitions for each level in this list
            foreach (var level in levels)
            {
                var levelDef = GetLevelDefinition(level, listBlocks);
                abstractNum.AppendChild(levelDef);
            }
            
            return abstractNum;
        }
        
        private Level GetLevelDefinition(int level, List<EnhancedBlock> listBlocks)
        {
            // Find a block with this level to determine format
            var sampleBlock = listBlocks.Find(b => b.Level == level);
            var numFmt = sampleBlock?.NumFmt ?? "decimal";
            var lvlText = GetLvlText(level, numFmt);
            
            var levelElement = new Level
            {
                LevelIndex = (Int32Value)level,
                TemplateCode = (HexBinaryValue)("0409" + level.ToString("X4")),
                StartNumberingValue = new StartNumberingValue { Val = (Int32Value)1 }
            };
            
            levelElement.AppendChild(new NumberingFormat { Val = GetNumberFormatValue(numFmt) });
            levelElement.AppendChild(new LevelText { Val = lvlText });
            levelElement.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
            
            // Add indentation
            var indent = new Indentation
            {
                Left = (StringValue)(level * 720).ToString(), // 720 twips = 0.5 inch
                Hanging = (StringValue)"360"
            };
            levelElement.AppendChild(indent);
            
            return levelElement;
        }
        
        private string GetLvlText(int level, string numFmt)
        {
            switch (numFmt)
            {
                case "decimal":
                    return level == 0 ? "%1." : "%1.%2.";
                case "upperLetter":
                    return "%1.";
                case "lowerLetter":
                    return "%1.";
                case "upperRoman":
                    return "%1.";
                case "lowerRoman":
                    return "%1.";
                default:
                    return "%1.";
            }
        }
        
        private NumberFormatValues GetNumberFormatValue(string numFmt)
        {
            return numFmt switch
            {
                "decimal" => NumberFormatValues.Decimal,
                "upperLetter" => NumberFormatValues.UpperLetter,
                "lowerLetter" => NumberFormatValues.LowerLetter,
                "upperRoman" => NumberFormatValues.UpperRoman,
                "lowerRoman" => NumberFormatValues.LowerRoman,
                _ => NumberFormatValues.Decimal
            };
        }
        
        private Paragraph CreateEnhancedParagraph(EnhancedBlock block, Dictionary<int, uint> numberingMap)
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties();
            
            // Add numbering properties if this is a list item
            if (block.IsListItem && block.Level.HasValue && block.ListId.HasValue)
            {
                var level = block.Level.Value;
                if (numberingMap.TryGetValue(level, out var numId))
                {
                    var numPr = new NumberingProperties();
                    numPr.AppendChild(new NumberingLevelReference { Val = (Int32Value)level });
                    numPr.AppendChild(new NumberingId { Val = (Int32Value)(int)numId });
                    pPr.AppendChild(numPr);
                    
                    Console.WriteLine($"Applied enhanced numbering level {level} with numId {numId} to: {block.CleanedContent}");
                }
                else
                {
                    Console.WriteLine($"No numbering found for level {level}: {block.CleanedContent}");
                }
            }
            else
            {
                Console.WriteLine($"Not a list item: {block.CleanedContent}");
            }
            
            paragraph.AppendChild(pPr);
            
            // Add the text run
            var run = new Run();
            var text = new Text(block.CleanedContent) { Space = SpaceProcessingModeValues.Preserve };
            run.AppendChild(text);
            paragraph.AppendChild(run);
            
            return paragraph;
        }
    }
    
    // class Program
    // {
    //     static void Main(string[] args)
    //     {
    //         if (args.Length != 3)
    //         {
    //             Console.WriteLine("Usage: EnhancedWordRebuilder.exe <enhanced_json> <template_docx> <output_docx>");
    //             return;
    //         }
    //         
    //         var jsonPath = args[0];
    //         var templatePath = args[1];
    //         var outputPath = args[2];
    //         
    //         var rebuilder = new EnhancedWordRebuilder();
    //         rebuilder.RebuildDocument(jsonPath, templatePath, outputPath);
    //     }
    // }
}
