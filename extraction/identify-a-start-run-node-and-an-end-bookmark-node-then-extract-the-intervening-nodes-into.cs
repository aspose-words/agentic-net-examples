using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class ExtractBetweenRunAndBookmark
{
    public static void Main()
    {
        // ---------- Create a sample source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Intro paragraph.");
        builder.Writeln("StartRun"); // Paragraph that contains the start Run.
        builder.Writeln("Content paragraph 1.");
        builder.Writeln("Content paragraph 2.");

        // Insert an end bookmark.
        builder.StartBookmark("EndBookmark");
        builder.Writeln("Inside bookmark.");
        builder.EndBookmark("EndBookmark");

        builder.Writeln("After bookmark.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ---------- Load the document for processing ----------
        Document loadedDoc = new Document(sourcePath);

        // Locate the start Run node with the exact text "StartRun".
        Run startRun = null;
        foreach (Run run in loadedDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "StartRun")
            {
                startRun = run;
                break;
            }
        }

        if (startRun == null)
            throw new InvalidOperationException("Start Run node not found.");

        // Locate the end bookmark.
        Bookmark endBookmark = loadedDoc.Range.Bookmarks["EndBookmark"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");

        // Determine the range of nodes between the start Run's paragraph and the bookmark start.
        Paragraph startParagraph = startRun.ParentParagraph;
        Node current = startParagraph.NextSibling;
        List<Node> nodesToExtract = new List<Node>();

        while (current != null && current != endBookmark.BookmarkStart)
        {
            nodesToExtract.Add(current);
            current = current.NextSibling;
        }

        if (nodesToExtract.Count == 0)
            throw new InvalidOperationException("No nodes found between the start Run and end bookmark.");

        // ---------- Create a new document to hold the extracted content ----------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in nodesToExtract)
        {
            Node importedNode = importer.ImportNode(node, true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed: output file was not created.");
    }
}
