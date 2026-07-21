using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a source document with a bookmark that encloses some nodes.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Source document - before bookmark.");
        srcBuilder.StartBookmark("ExtractHere");
        srcBuilder.Writeln("First paragraph inside bookmark.");
        srcBuilder.Writeln("Second paragraph inside bookmark.");
        srcBuilder.EndBookmark("ExtractHere");
        srcBuilder.Writeln("Source document - after bookmark.");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and extract the nodes that are inside the bookmark.
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);
        Bookmark sourceBookmark = loadedSource.Range.Bookmarks["ExtractHere"];
        if (sourceBookmark == null)
            throw new InvalidOperationException("Source bookmark not found.");

        // The bookmark start and end nodes act as markers; the actual content lies between them.
        Node startNode = sourceBookmark.BookmarkStart;
        Node endNode = sourceBookmark.BookmarkEnd;

        List<Node> extractedNodes = new List<Node>();
        Node cur = startNode.NextSibling;
        while (cur != null && cur != endNode)
        {
            extractedNodes.Add(cur);
            cur = cur.NextSibling;
        }

        if (extractedNodes.Count == 0)
            throw new InvalidOperationException("No nodes were extracted from the source bookmark.");

        // -----------------------------------------------------------------
        // 3. Create a destination document that contains a custom bookmark where the extracted nodes will be inserted.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("Destination document - before insertion point.");
        destBuilder.StartBookmark("InsertPoint");
        // The bookmark is intentionally left empty.
        destBuilder.EndBookmark("InsertPoint");
        destBuilder.Writeln("Destination document - after insertion point.");

        const string destPath = "destination.docx";
        destDoc.Save(destPath);

        // -----------------------------------------------------------------
        // 4. Load the destination document and insert the extracted nodes at the custom bookmark.
        // -----------------------------------------------------------------
        Document loadedDest = new Document(destPath);
        Bookmark destBookmark = loadedDest.Range.Bookmarks["InsertPoint"];
        if (destBookmark == null)
            throw new InvalidOperationException("Destination bookmark not found.");

        // The insertion point is the bookmark start node.
        Node insertionNode = destBookmark.BookmarkStart;
        CompositeNode parent = insertionNode.ParentNode;

        // Use NodeImporter to import nodes from the source document into the destination document.
        NodeImporter importer = new NodeImporter(loadedSource, loadedDest, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in extractedNodes)
        {
            Node importedNode = importer.ImportNode(node, true);
            parent.InsertAfter(importedNode, insertionNode);
            insertionNode = importedNode; // Move the insertion point forward.
        }

        const string resultPath = "result.docx";
        loadedDest.Save(resultPath);

        // -----------------------------------------------------------------
        // 5. Validate that the result file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        // Example completed successfully.
    }
}
