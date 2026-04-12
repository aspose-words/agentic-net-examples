using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------- Create source document with a bookmark --------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Paragraph before bookmark.");

        // Insert bookmark start.
        srcBuilder.StartBookmark("SourceBookmark");
        srcBuilder.Writeln("First paragraph inside bookmark.");
        srcBuilder.Writeln("Second paragraph inside bookmark.");
        // Insert bookmark end.
        srcBuilder.EndBookmark("SourceBookmark");

        srcBuilder.Writeln("Paragraph after bookmark.");

        // Save the source document (optional, for verification).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------- Extract nodes that are inside the bookmark --------------------
        Bookmark sourceBookmark = sourceDoc.Range.Bookmarks["SourceBookmark"];
        if (sourceBookmark == null)
            throw new InvalidOperationException("Source bookmark not found.");

        BookmarkStart bookmarkStart = sourceBookmark.BookmarkStart;
        BookmarkEnd bookmarkEnd = sourceBookmark.BookmarkEnd;

        List<Node> extractedNodes = new List<Node>();
        Node currentNode = bookmarkStart.NextSibling;
        while (currentNode != null && currentNode != bookmarkEnd)
        {
            extractedNodes.Add(currentNode);
            currentNode = currentNode.NextSibling;
        }

        if (extractedNodes.Count == 0)
            throw new InvalidOperationException("No nodes were extracted from the bookmark.");

        // -------------------- Create destination document with insertion bookmark --------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Document header.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Placeholder paragraph before insertion.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Document footer.");

        // Move the builder to the insertion bookmark.
        destBuilder.MoveToBookmark("InsertHere");

        // -------------------- Insert extracted nodes into destination document --------------------
        foreach (Node node in extractedNodes)
        {
            // Import the node into the destination document to maintain correct document relationships.
            Node importedNode = destDoc.ImportNode(node, true, ImportFormatMode.KeepSourceFormatting);
            destBuilder.InsertNode(importedNode);
        }

        // -------------------- Save destination document --------------------
        string destPath = "Destination.docx";
        destDoc.Save(destPath);

        // Validate that the destination file was created.
        if (!File.Exists(destPath))
            throw new InvalidOperationException("Failed to create the destination document.");

        Console.WriteLine("Extraction and insertion completed successfully.");
    }
}
