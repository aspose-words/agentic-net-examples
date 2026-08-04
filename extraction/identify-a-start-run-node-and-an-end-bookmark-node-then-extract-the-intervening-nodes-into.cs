using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with a start Run and an end Bookmark.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the range.
        builder.Writeln("Paragraph before the range.");

        // Insert the start Run that will mark the beginning of extraction.
        Run startRun = new Run(sourceDoc, "StartRun");
        builder.CurrentParagraph.AppendChild(startRun);
        builder.Writeln(); // End the paragraph containing the start run.

        // Content that should be extracted (paragraphs and a table).
        builder.Writeln("First extracted paragraph.");
        builder.Writeln("Second extracted paragraph.");

        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Insert the end Bookmark that will mark the end of extraction.
        builder.StartBookmark("EndMarker");
        builder.Writeln("Paragraph after the range (inside bookmark).");
        builder.EndBookmark("EndMarker");

        // Paragraph after the range.
        builder.Writeln("Paragraph after the range.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and locate the start Run and end Bookmark.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Find the start Run by its exact text.
        Run foundStartRun = null;
        foreach (Run run in loadedDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "StartRun")
            {
                foundStartRun = run;
                break;
            }
        }

        if (foundStartRun == null)
            throw new InvalidOperationException("Start Run not found.");

        // Find the end Bookmark by name.
        Bookmark endBookmark = loadedDoc.Range.Bookmarks["EndMarker"];
        if (endBookmark == null)
            throw new InvalidOperationException("End Bookmark not found.");

        // Use the BookmarkStart node as the exclusive end boundary.
        Node endNode = endBookmark.BookmarkStart;

        // -----------------------------------------------------------------
        // 3. Collect all block‑level nodes that lie between the start Run and the end Bookmark.
        // -----------------------------------------------------------------
        List<Node> nodesToExtract = new List<Node>();

        // The start Run resides inside its own paragraph. We begin extraction after that paragraph.
        Node current = foundStartRun.ParentNode?.NextSibling;

        while (current != null && !current.Equals(endNode))
        {
            if (current.NodeType == NodeType.Paragraph || current.NodeType == NodeType.Table)
                nodesToExtract.Add(current);

            current = current.NextSibling;
        }

        if (nodesToExtract.Count == 0)
            throw new InvalidOperationException("No nodes were found between the start Run and the end Bookmark.");

        // -----------------------------------------------------------------
        // 4. Create a new document and import the collected nodes.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in nodesToExtract)
        {
            Node importedNode = importer.ImportNode(node, true);
            resultBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 5. Save the extracted content and verify the output file.
        // -----------------------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
