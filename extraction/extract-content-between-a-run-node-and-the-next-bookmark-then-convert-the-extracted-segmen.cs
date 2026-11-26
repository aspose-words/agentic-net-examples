using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class ExtractRunToBookmark
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a sample document with runs and a bookmark.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the target area.
        builder.Writeln("Paragraph before the target area.");

        // First run – this will be the start node for extraction.
        builder.Write("RunA");

        // Additional run that lies between the first run and the bookmark.
        builder.Write("BetweenRun");

        // Insert a bookmark after the runs.
        builder.StartBookmark("TargetBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("TargetBookmark");

        // Paragraph after the bookmark.
        builder.Writeln("Paragraph after the bookmark.");

        // Save the source document (optional, for inspection).
        sourceDoc.Save("source.docx");

        // -------------------------------------------------
        // Locate the start run (the run containing "RunA").
        // -------------------------------------------------
        Run startRun = null;
        foreach (Run run in sourceDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "RunA")
            {
                startRun = run;
                break;
            }
        }

        if (startRun == null)
            throw new InvalidOperationException("Start run with text 'RunA' not found.");

        // -------------------------------------------------
        // Find the next bookmark start after the run.
        // -------------------------------------------------
        BookmarkStart nextBookmarkStart = null;
        Node current = startRun.NextSibling;
        while (current != null && nextBookmarkStart == null)
        {
            if (current.NodeType == NodeType.BookmarkStart)
                nextBookmarkStart = (BookmarkStart)current;
            else
                current = current.NextSibling;
        }

        if (nextBookmarkStart == null)
            throw new InvalidOperationException("Next bookmark start not found after the run.");

        // -------------------------------------------------
        // Collect nodes that lie between the run and the bookmark (exclusive).
        // -------------------------------------------------
        List<Node> nodesBetween = new List<Node>();
        Node nodeBetween = startRun.NextSibling;
        while (nodeBetween != null && nodeBetween != nextBookmarkStart)
        {
            nodesBetween.Add(nodeBetween);
            nodeBetween = nodeBetween.NextSibling;
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No nodes found between the run and the bookmark.");

        // -------------------------------------------------
        // Build a new document containing the extracted nodes.
        // -------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Create a new section and body for the destination document.
        Section destSection = new Section(extractedDoc);
        extractedDoc.AppendChild(destSection);
        Body destBody = new Body(extractedDoc);
        destSection.AppendChild(destBody);

        // The extracted nodes are inline (runs), so we need a paragraph container.
        Paragraph destParagraph = new Paragraph(extractedDoc);
        destBody.AppendChild(destParagraph);

        // Use NodeImporter to import nodes from sourceDoc into extractedDoc.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node srcNode in nodesBetween)
        {
            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(srcNode, true);
            // The imported node will be a Run (or other inline node). Append it to the paragraph.
            destParagraph.AppendChild(importedNode);
        }

        // -------------------------------------------------
        // Save the extracted segment as HTML.
        // -------------------------------------------------
        string htmlPath = "extracted.html";
        extractedDoc.Save(htmlPath, SaveFormat.Html);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output was not created.");

        Console.WriteLine("Extraction and HTML conversion completed successfully.");
    }
}
