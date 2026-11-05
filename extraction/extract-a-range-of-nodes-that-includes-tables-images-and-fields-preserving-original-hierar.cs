using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class ExtractionExample
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with various nodes inside a bookmark.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Document header - not part of extraction.");

        builder.StartBookmark("ExtractRange");

        builder.Writeln("This paragraph is inside the extraction range.");

        // Table inside the range.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Image inside the range.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Field (DATE) inside the range.
        builder.InsertField("DATE \\@ \"MMMM d, yyyy\"");

        builder.EndBookmark("ExtractRange");

        builder.Writeln("Document footer - not part of extraction.");

        sourceDoc.Save("source.docx");

        // -------------------------------------------------
        // 2. Load the document and locate the bookmark that defines the extraction range.
        // -------------------------------------------------
        Document loadedDoc = new Document("source.docx");
        Bookmark extractBookmark = loadedDoc.Range.Bookmarks["ExtractRange"];
        if (extractBookmark == null)
            throw new InvalidOperationException("Required bookmark 'ExtractRange' was not found.");

        // -------------------------------------------------
        // 3. Prepare a new document that will contain the extracted nodes.
        // -------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section newSection = new Section(extractedDoc);
        extractedDoc.AppendChild(newSection);

        Body newBody = new Body(extractedDoc);
        newSection.AppendChild(newBody);

        // -------------------------------------------------
        // 4. Clone (import) block-level nodes that lie between the bookmark start and end.
        // -------------------------------------------------
        Node startNode = extractBookmark.BookmarkStart;
        Node endNode = extractBookmark.BookmarkEnd;

        Body sourceBody = loadedDoc.FirstSection.Body;
        bool insideRange = false;

        // Importer to move nodes from source to destination document.
        NodeImporter importer = new NodeImporter(loadedDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in sourceBody.GetChildNodes(NodeType.Any, true))
        {
            if (node == startNode)
            {
                insideRange = true;
                continue; // Skip the bookmark start node itself.
            }

            if (node == endNode)
            {
                break; // Stop after reaching the bookmark end node.
            }

            if (!insideRange)
                continue;

            // Only block-level nodes (Paragraph, Table) can be direct children of Body.
            if (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table)
            {
                Node importedNode = importer.ImportNode(node, true);
                newBody.AppendChild(importedNode);
            }
        }

        // -------------------------------------------------
        // 5. Save the extracted document.
        // -------------------------------------------------
        extractedDoc.Save("extracted.docx");
        if (!File.Exists("extracted.docx"))
            throw new InvalidOperationException("Extraction failed: output file not created.");

        // -------------------------------------------------
        // 6. Generate a JSON report with counts of tables, images, and fields.
        // -------------------------------------------------
        var report = new
        {
            Tables = extractedDoc.GetChildNodes(NodeType.Table, true).Count,
            Images = extractedDoc.GetChildNodes(NodeType.Shape, true)
                                 .OfType<Shape>()
                                 .Count(s => s.HasImage),
            Fields = extractedDoc.GetChildNodes(NodeType.FieldStart, true).Count
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("extraction-report.json", jsonReport);
        if (!File.Exists("extraction-report.json"))
            throw new InvalidOperationException("Report generation failed: JSON file not created.");
    }
}
