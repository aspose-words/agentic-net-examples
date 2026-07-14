using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Newtonsoft.Json;

public class ExtractionExample
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document containing a table, an image,
        //    and a field, wrapped between two bookmarks.
        // -----------------------------------------------------------------
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Document introduction paragraph.");

        // Start bookmark.
        builder.StartBookmark("Start");
        builder.EndBookmark("Start");

        // Insert a simple 1x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Insert a tiny PNG image.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (var imageStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imageStream);
        }

        // Insert a MERGEFIELD.
        builder.InsertField("MERGEFIELD SampleField \\* MERGEFORMAT");

        // End bookmark.
        builder.StartBookmark("End");
        builder.EndBookmark("End");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and locate the start/end bookmarks.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(sourcePath);
        var startBookmark = loadedDoc.Range.Bookmarks["Start"];
        var endBookmark = loadedDoc.Range.Bookmarks["End"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // -----------------------------------------------------------------
        // 3. Prepare the destination document (empty).
        // -----------------------------------------------------------------
        var resultDoc = new Document();
        resultDoc.RemoveAllChildren();
        var resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        var resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // 4. Extract nodes that lie between the two bookmarks (exclusive).
        //    Preserve the original hierarchy by importing each node.
        // -----------------------------------------------------------------
        var importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node currentNode = startBookmark.BookmarkStart.NextSibling;
        while (currentNode != null && currentNode != endBookmark.BookmarkEnd)
        {
            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(currentNode, true);
            resultBody.AppendChild(importedNode);
            currentNode = currentNode.NextSibling;
        }

        // -----------------------------------------------------------------
        // 5. Save the extracted content.
        // -----------------------------------------------------------------
        const string extractedPath = "extracted.docx";
        resultDoc.Save(extractedPath);
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("Extraction output file was not created.");

        // -----------------------------------------------------------------
        // 6. Build a simple JSON report about the extracted content.
        // -----------------------------------------------------------------
        var report = new
        {
            ParagraphCount = resultDoc.FirstSection.Body.Paragraphs.Count,
            TableCount = resultDoc.GetChildNodes(NodeType.Table, true).Count,
            ImageCount = resultDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Count(s => s.HasImage),
            FieldCount = resultDoc.Range.Fields.Count
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string reportPath = "extraction-report.json";
        File.WriteAllText(reportPath, jsonReport);
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Extraction report file was not created.");

        // -----------------------------------------------------------------
        // 7. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine("Extraction completed successfully.");
        Console.WriteLine($"Extracted document: {extractedPath}");
        Console.WriteLine($"Report file: {reportPath}");
    }
}
