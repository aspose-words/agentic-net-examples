using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document that contains a bookmark range
        //    with a paragraph, a table, an image and a field.
        // -----------------------------------------------------------------
        string sourcePath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Document header - not part of extraction.");

        // Bookmark that defines the extraction range.
        builder.StartBookmark("ExtractStart");
        builder.Writeln("Paragraph inside extraction range.");

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
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=");
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        // Field inside the range.
        builder.InsertField("MERGEFIELD SampleField \\* MERGEFORMAT");
        builder.EndBookmark("ExtractStart");

        builder.Writeln("Document footer - not part of extraction.");

        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and locate the bookmark boundaries.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        Bookmark startBookmark = loadedDoc.Range.Bookmarks["ExtractStart"];
        if (startBookmark == null)
            throw new InvalidOperationException("Start bookmark not found.");

        BookmarkEnd endBookmarkNode = startBookmark.BookmarkEnd;
        if (endBookmarkNode == null)
            throw new InvalidOperationException("End bookmark node not found.");

        Node startNode = startBookmark.BookmarkStart.ParentNode;
        Node endNode = endBookmarkNode.ParentNode;
        if (startNode == null || endNode == null)
            throw new InvalidOperationException("Unable to locate range boundaries.");

        // -----------------------------------------------------------------
        // 3. Collect all block-level nodes that lie between the start and end
        //    nodes (inclusive).
        // -----------------------------------------------------------------
        List<Node> nodesToExtract = new List<Node>();
        for (Node cur = startNode; cur != null; cur = cur.NextSibling)
        {
            nodesToExtract.Add(cur);
            if (cur == endNode)
                break;
        }

        // -----------------------------------------------------------------
        // 4. Create a new document that will hold the extracted content.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -----------------------------------------------------------------
        // 5. Import each collected node into the destination document.
        //    NodeImporter handles the document ownership conversion.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in nodesToExtract)
        {
            Node importedNode = importer.ImportNode(node, true);
            if (importedNode != null)
                resultBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 6. Save the extracted range.
        // -----------------------------------------------------------------
        string outputPath = "extracted.docx";
        resultDoc.Save(outputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 7. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Extraction output file was not created.");
    }
}
