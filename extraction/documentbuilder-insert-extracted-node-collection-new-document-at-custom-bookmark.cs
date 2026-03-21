using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a source document with sample content.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("First paragraph from source.");
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndTable();
        srcBuilder.Writeln("Second paragraph from source.");

        // Create a new destination document.
        Document dstDoc = new Document();

        // Add a bookmark in the destination document where the nodes will be inserted.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.StartBookmark("InsertHere");
        builder.Writeln("Text before insertion.");
        builder.EndBookmark("InsertHere");

        // Collect all block‑level nodes from the source document's first section body.
        List<Node> nodesToInsert = srcDoc.Sections[0].Body.GetChildNodes(NodeType.Any, true)
            .Cast<Node>()
            .Where(n => n.NodeType == NodeType.Paragraph || n.NodeType == NodeType.Table)
            .ToList();

        // Insert the collected nodes at the bookmark location.
        InsertNodeCollectionAtBookmark(dstDoc, srcDoc, nodesToInsert, "InsertHere");

        // Save the resulting document.
        dstDoc.Save("Result.docx");
        Console.WriteLine("Result.docx created successfully.");
    }

    static void InsertNodeCollectionAtBookmark(Document dstDoc, Document srcDoc, List<Node> nodes, string bookmarkName)
    {
        // Move the builder cursor to the start of the bookmark.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        if (!builder.MoveToBookmark(bookmarkName))
            throw new ArgumentException($"Bookmark '{bookmarkName}' not found in the destination document.");

        // Determine the node after which we will start inserting.
        Node insertionPoint = builder.CurrentParagraph ?? builder.CurrentNode;
        if (insertionPoint == null)
            insertionPoint = dstDoc.Range.Bookmarks[bookmarkName].BookmarkStart;

        // Create a NodeImporter for efficient repeated imports.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert each node after the previous insertion point.
        foreach (Node srcNode in nodes)
        {
            // Skip the last empty paragraph of a section (mirrors Aspose example logic).
            if (srcNode.NodeType == NodeType.Paragraph)
            {
                Paragraph para = (Paragraph)srcNode;
                if (para.IsEndOfSection && !para.HasChildNodes)
                    continue;
            }

            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(srcNode, true);

            // Insert the imported node after the current insertion point.
            CompositeNode parent = insertionPoint.ParentNode as CompositeNode;
            parent.InsertAfter(importedNode, insertionPoint);

            // Update the insertion point for the next iteration.
            insertionPoint = importedNode;
        }
    }
}
