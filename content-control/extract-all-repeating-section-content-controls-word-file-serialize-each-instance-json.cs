using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using System.Text.Json;

namespace AsposeWordsRepeatingSectionExport
{
    // Model that represents a repeating section content control.
    public class RepeatingSectionInfo
    {
        public int Id { get; set; }               // Identifier of the content control.
        public string Title { get; set; }          // Title (if any) of the content control.
        public List<string> Items { get; set; }    // Text of each repeating item.
    }

    public class Program
    {
        // Entry point.
        public static void Main(string[] args)
        {
            // Example usage:
            // args[0] – path to the source .docx file.
            // args[1] – path where the resulting JSON will be written.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeWordsRepeatingSectionExport <input.docx> <output.json>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                var sections = ExtractRepeatingSections(inputPath);
                // Use System.Text.Json for serialization (built‑in, no external package required).
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(sections, jsonOptions);
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Successfully extracted {sections.Count} repeating sections to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Loads the document, finds all repeating section SDTs and builds a list of serializable objects.
        private static List<RepeatingSectionInfo> ExtractRepeatingSections(string docPath)
        {
            // Load the Word document (lifecycle rule: use the Document constructor for loading).
            Document doc = new Document(docPath);

            // Collect all StructuredDocumentTag nodes in the document.
            NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            var result = new List<RepeatingSectionInfo>();

            foreach (StructuredDocumentTag sdt in sdtNodes)
            {
                // We are interested only in repeating section content controls.
                if (sdt.SdtType != SdtType.RepeatingSection)
                    continue;

                var info = new RepeatingSectionInfo
                {
                    Id = sdt.Id,
                    Title = sdt.Title,
                    Items = new List<string>()
                };

                // A repeating section contains child SDTs of type RepeatingSectionItem.
                // Each item represents one instance of the repeated content.
                NodeCollection itemNodes = sdt.GetChildNodes(NodeType.StructuredDocumentTag, true);
                foreach (StructuredDocumentTag item in itemNodes)
                {
                    if (item.SdtType != SdtType.RepeatingSectionItem)
                        continue;

                    // Retrieve the textual representation of the item's content.
                    // Using GetText() gives the plain text inside the item.
                    string itemText = item.GetText().Trim();
                    info.Items.Add(itemText);
                }

                result.Add(info);
            }

            return result;
        }
    }
}
