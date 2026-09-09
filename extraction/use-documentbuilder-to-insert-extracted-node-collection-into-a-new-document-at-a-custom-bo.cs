using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with a bookmark that encloses several nodes.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Paragraph before bookmark.");

        srcBuilder.StartBookmark("ExtractMe");
        srcBuilder.Writeln("First paragraph inside bookmark.");
        srcBuilder.Writeln("Second paragraph inside bookmark.");

        // Insert a simple table inside the bookmark.
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcBuilder.Writeln("Third paragraph inside bookmark.");
        srcBuilder.EndBookmark("ExtractMe");

        srcBuilder.Writeln("Paragraph after bookmark.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedSource = new Document(sourcePath);

        // Retrieve the bookmark that defines the range to extract.
        Bookmark extractBookmark = loadedSource.Range.Bookmarks["ExtractMe"];
        if (extractBookmark == null)
            throw new InvalidOperationException("Bookmark 'ExtractMe' was not found in the source document.");

        // Collect all nodes that are directly between the bookmark start and end.
        List<Node> extractedNodes = new List<Node>();
        Node current = extractBookmark.BookmarkStart.NextSibling;
        while (current != null && current != extractBookmark.BookmarkEnd)
        {
            extractedNodes.Add(current);
            current = current.NextSibling;
        }

        if (extractedNodes.Count == 0)
            throw new InvalidOperationException("No nodes were extracted from the bookmark.");

        // Create a destination document with a custom bookmark where the extracted nodes will be inserted.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Content before insertion point.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Placeholder paragraph.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Content after insertion point.");

        // Move to the insertion bookmark.
        Bookmark insertBookmark = destDoc.Range.Bookmarks["InsertHere"];
        if (insertBookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' was not found in the destination document.");

        // Prepare a NodeImporter for importing nodes from the source to the destination.
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert each extracted node after the bookmark start node.
        Node insertionPoint = insertBookmark.BookmarkStart;
        foreach (Node node in extractedNodes)
        {
            Node importedNode = importer.ImportNode(node, true);
            insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
            insertionPoint = importedNode;
        }

        // Save the destination document.
        const string resultPath = "result.docx";
        destDoc.Save(resultPath);

        // Verify that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        // Optional: indicate success (no console interaction required by the task).
    }
}
