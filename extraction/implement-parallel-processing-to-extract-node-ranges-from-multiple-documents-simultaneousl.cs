using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "Extracted");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create sample documents that contain two bookmarks ("Start" and "End")
        // surrounding the paragraphs we want to extract.
        // -----------------------------------------------------------------
        const int docCount = 5;
        for (int i = 1; i <= docCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Intro paragraph.");
            builder.StartBookmark("Start");
            builder.Writeln($"Document {i} - First extracted paragraph.");
            builder.Writeln($"Document {i} - Second extracted paragraph.");
            builder.EndBookmark("Start");
            builder.StartBookmark("End");
            builder.Writeln($"Document {i} - Content after extraction.");
            builder.EndBookmark("End");
            builder.Writeln($"Document {i} - Final paragraph.");

            string filePath = Path.Combine(inputDir, $"doc{i}.docx");
            doc.Save(filePath);
        }

        // -----------------------------------------------------------------
        // Process all documents in parallel, extracting the content that lies
        // between the "Start" and "End" bookmarks (exclusive) and saving it
        // as a new document.
        // -----------------------------------------------------------------
        string[] files = Directory.GetFiles(inputDir, "*.docx");
        var summary = new List<object>();

        Parallel.ForEach(files, filePath =>
        {
            // Load the source document.
            Document source = new Document(filePath);

            // Retrieve the bookmarks that define the extraction range.
            Bookmark startBookmark = source.Range.Bookmarks["Start"];
            Bookmark endBookmark = source.Range.Bookmarks["End"];
            if (startBookmark == null || endBookmark == null)
                throw new InvalidOperationException($"Bookmarks not found in {Path.GetFileName(filePath)}.");

            // Create a fresh destination document with a single section and body.
            Document extracted = new Document();
            extracted.RemoveAllChildren(); // Remove the default section/body.
            Section destSection = new Section(extracted);
            extracted.AppendChild(destSection);
            Body destBody = new Body(extracted);
            destSection.AppendChild(destBody);

            // Prepare an importer to copy nodes from the source to the destination.
            NodeImporter importer = new NodeImporter(source, extracted, ImportFormatMode.KeepSourceFormatting);

            // Walk through the nodes that lie between the two bookmarks.
            Node current = startBookmark.BookmarkStart.NextSibling;
            while (current != null && current != endBookmark.BookmarkEnd)
            {
                // Only block-level nodes (Paragraph, Table, etc.) can be appended to a Body.
                if (current.NodeType == NodeType.Paragraph || current.NodeType == NodeType.Table)
                {
                    Node importedNode = importer.ImportNode(current, true);
                    destBody.AppendChild(importedNode);
                }

                current = current.NextSibling;
            }

            // Save the extracted document.
            string outputFileName = $"extracted-{Path.GetFileNameWithoutExtension(filePath)}.docx";
            string outputPath = Path.Combine(outputDir, outputFileName);
            extracted.Save(outputPath);

            // Record the processing result.
            lock (summary)
            {
                summary.Add(new { Input = Path.GetFileName(filePath), Output = outputFileName });
            }
        });

        // -----------------------------------------------------------------
        // Verify that every expected output file exists.
        // -----------------------------------------------------------------
        foreach (var entry in summary)
        {
            string outputPath = Path.Combine(outputDir, ((dynamic)entry).Output);
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Expected output file was not created: {outputPath}");
        }

        // Write a JSON report describing the batch operation.
        string jsonReportPath = Path.Combine(baseDir, "extraction-report.json");
        File.WriteAllText(jsonReportPath, JsonConvert.SerializeObject(summary, Formatting.Indented));

        // Final validation of the JSON report.
        if (!File.Exists(jsonReportPath))
            throw new InvalidOperationException("JSON report was not created.");
    }
}
