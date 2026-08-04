using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with various content.
        // -------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Intro paragraph (outside the extraction range).
        builder.Writeln("Intro paragraph before the range.");

        // Bookmark that marks the start of the extraction range.
        builder.StartBookmark("ExtractStart");
        builder.Writeln("Paragraph inside the range - before table.");

        // Insert a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Insert a 1x1 pixel PNG image from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Insert a DATE field.
        builder.InsertField(@" DATE \@ ""MMMM d, yyyy"" ");

        builder.Writeln("Paragraph inside the range - after field.");
        // End bookmark that marks the end of the extraction range.
        builder.EndBookmark("ExtractStart");

        // Paragraph after the range.
        builder.Writeln("Paragraph after the range.");

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the source document for extraction.
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        // Retrieve the bookmark that defines the extraction boundaries.
        Bookmark bookmark = loaded.Range.Bookmarks["ExtractStart"];
        if (bookmark == null)
            throw new InvalidOperationException("Extraction bookmark not found.");

        // Determine the first and last block-level nodes that belong to the bookmarked range.
        Node startNode = bookmark.BookmarkStart.ParentNode; // Usually a Paragraph.
        Node endNode = bookmark.BookmarkEnd.ParentNode;     // Usually a Paragraph.

        // Collect all block-level nodes between startNode and endNode inclusive.
        var nodesToExtract = new List<Node>();
        Node current = startNode;
        while (current != null)
        {
            nodesToExtract.Add(current);
            if (current == endNode)
                break;
            current = current.NextSibling;
        }

        // -------------------------------------------------
        // 3. Build the destination document and import the collected nodes.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren(); // Ensure a clean document.

        // Create a new section with a body.
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes while preserving formatting.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in nodesToExtract)
        {
            // Import the node (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(node, true);
            resultBody.AppendChild(importedNode);
        }

        // -------------------------------------------------
        // 4. Save the extracted range to a new document.
        // -------------------------------------------------
        const string resultPath = "extracted_range.docx";
        result.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        Console.WriteLine($"Extraction completed. Output saved to '{resultPath}'.");
    }
}
