using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Required by Aspose.Words for image handling
using Newtonsoft.Json;

namespace ExtractionExample
{
    // Reusable extraction utility
    public static class ExtractionUtility
    {
        // Extracts content based on a simple identifier and returns a new Document.
        // Identifier formats:
        //   "Paragraph:{zeroBasedIndex}"
        //   "Table:{zeroBasedIndex}"
        //   "Bookmark:{bookmarkName}"
        public static Document ExtractContent(Document source, string identifier)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(identifier)) throw new ArgumentException("Identifier cannot be null or empty.", nameof(identifier));

            // Prepare a clean destination document.
            Document result = new Document();
            result.RemoveAllChildren();
            Section section = new Section(result);
            result.AppendChild(section);
            Body body = new Body(result);
            section.AppendChild(body);

            if (identifier.StartsWith("Paragraph:", StringComparison.OrdinalIgnoreCase))
            {
                // Extract a specific paragraph.
                string indexPart = identifier.Substring("Paragraph:".Length);
                if (!int.TryParse(indexPart, out int paraIndex))
                    throw new ArgumentException("Invalid paragraph index.", nameof(identifier));

                Paragraph paragraph = source.FirstSection?.Body?.Paragraphs[paraIndex];
                if (paragraph == null)
                    throw new InvalidOperationException($"Paragraph at index {paraIndex} not found.");

                // Import the cloned paragraph into the result document.
                Node imported = result.ImportNode(paragraph.Clone(true), true, ImportFormatMode.KeepSourceFormatting);
                body.AppendChild(imported);
            }
            else if (identifier.StartsWith("Table:", StringComparison.OrdinalIgnoreCase))
            {
                // Extract a specific table.
                string indexPart = identifier.Substring("Table:".Length);
                if (!int.TryParse(indexPart, out int tableIndex))
                    throw new ArgumentException("Invalid table index.", nameof(identifier));

                NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
                if (tableIndex < 0 || tableIndex >= tables.Count)
                    throw new InvalidOperationException($"Table at index {tableIndex} not found.");

                Table table = tables[tableIndex] as Table;
                if (table == null)
                    throw new InvalidOperationException("Failed to cast node to Table.");

                // Import the cloned table into the result document.
                Node imported = result.ImportNode(table.Clone(true), true, ImportFormatMode.KeepSourceFormatting);
                body.AppendChild(imported);
            }
            else if (identifier.StartsWith("Bookmark:", StringComparison.OrdinalIgnoreCase))
            {
                // Extract the text of a bookmark.
                string bookmarkName = identifier.Substring("Bookmark:".Length);
                Bookmark bookmark = source.Range.Bookmarks[bookmarkName];
                if (bookmark == null)
                    throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

                // Create a paragraph containing the bookmark's text in the result document.
                Paragraph para = new Paragraph(result);
                Run run = new Run(result, bookmark.Text);
                para.AppendChild(run);
                body.AppendChild(para);
            }
            else
            {
                throw new ArgumentException("Unsupported identifier format.", nameof(identifier));
            }

            return result;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a sample source document.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Add paragraphs.
            builder.Writeln("Paragraph 0 - introductory text.");
            builder.Writeln("Paragraph 1 - contains a bookmark.");
            builder.StartBookmark("SampleBookmark");
            builder.Writeln("This is the bookmarked text.");
            builder.EndBookmark("SampleBookmark");
            builder.Writeln("Paragraph 2 - after bookmark.");

            // Add a table.
            builder.StartTable();
            builder.InsertCell(); builder.Write("Cell A1");
            builder.InsertCell(); builder.Write("Cell B1");
            builder.EndRow();
            builder.InsertCell(); builder.Write("Cell A2");
            builder.InsertCell(); builder.Write("Cell B2");
            builder.EndRow();
            builder.EndTable();

            // Save the source document.
            const string sourcePath = "sample.docx";
            sourceDoc.Save(sourcePath);

            // Verify source file creation.
            if (!File.Exists(sourcePath))
                throw new InvalidOperationException("Source document was not created.");

            // Step 2: Load the source document.
            Document loadedDoc = new Document(sourcePath);

            // Step 3: Perform extractions using the utility.
            var extractions = new Dictionary<string, string>
            {
                { "Paragraph:1", "extracted-paragraph.docx" },
                { "Table:0", "extracted-table.docx" },
                { "Bookmark:SampleBookmark", "extracted-bookmark.docx" }
            };

            foreach (var kvp in extractions)
            {
                string identifier = kvp.Key;
                string outputPath = kvp.Value;

                Document extracted = ExtractionUtility.ExtractContent(loadedDoc, identifier);
                extracted.Save(outputPath);

                // Validate output.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException($"Extraction failed for identifier '{identifier}'. Output file not found.");
            }

            // Optional: Serialize extraction metadata to JSON.
            var metadata = new
            {
                SourceDocument = sourcePath,
                Extractions = extractions
            };
            string jsonPath = "extraction-metadata.json";
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(metadata, Formatting.Indented));

            // Validate JSON output.
            if (!File.Exists(jsonPath))
                throw new InvalidOperationException("Metadata JSON file was not created.");

            // Program completed successfully.
        }
    }
}
