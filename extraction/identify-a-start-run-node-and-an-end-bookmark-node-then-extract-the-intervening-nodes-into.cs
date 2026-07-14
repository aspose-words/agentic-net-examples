using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Text before the start run.
        builder.Writeln("Paragraph before start run.");

        // Insert a paragraph containing the start run.
        Paragraph startParagraph = new Paragraph(source);
        Run startRun = new Run(source, "StartRun");
        startParagraph.AppendChild(startRun);
        source.FirstSection.Body.AppendChild(startParagraph);

        // Add some content that will be extracted.
        builder.Writeln("First extracted paragraph.");
        builder.Writeln("Second extracted paragraph.");

        // Insert a table that will also be extracted.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Insert the end bookmark.
        builder.StartBookmark("EndBookmark");
        builder.Writeln("Paragraph inside end bookmark.");
        builder.EndBookmark("EndBookmark");

        // Text after the end bookmark.
        builder.Writeln("Paragraph after end bookmark.");

        // Save the source document.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // Load the document for processing.
        Document doc = new Document(sourcePath);

        // Locate the start run node by its text.
        Run startRunNode = null;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "StartRun")
            {
                startRunNode = run;
                break;
            }
        }
        if (startRunNode == null)
            throw new InvalidOperationException("Start run node not found.");

        // Locate the end bookmark.
        Bookmark endBookmark = doc.Range.Bookmarks["EndBookmark"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");

        // The bookmark start node marks the end boundary.
        Node endNode = endBookmark.BookmarkStart;

        // Collect nodes that lie between the start run and the end bookmark (exclusive).
        List<Node> extractedNodes = new List<Node>();
        Node current = startRunNode;
        while (true)
        {
            current = current.NextPreOrder(doc);
            if (current == null || current == endNode)
                break;

            // Skip the start run itself; we want intervening nodes only.
            extractedNodes.Add(current);
        }

        // Prepare the destination document.
        Document result = new Document();
        result.RemoveAllChildren();
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Import and append the collected nodes.
        NodeImporter importer = new NodeImporter(doc, result, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in extractedNodes)
        {
            // Only block-level nodes can be appended directly to the body.
            if (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table)
            {
                Node imported = importer.ImportNode(node, true);
                resultBody.AppendChild(imported);
            }
            else if (node.NodeType == NodeType.Run)
            {
                // Wrap inline runs in a new paragraph.
                Paragraph para = new Paragraph(result);
                Node importedRun = importer.ImportNode(node, true);
                para.AppendChild(importedRun);
                resultBody.AppendChild(para);
            }
            // Other node types are ignored for this example.
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
