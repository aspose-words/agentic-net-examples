using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document that contains the nodes we want to copy.
        Document srcDoc = new Document("Source.docx");

        // Create a new destination document.
        Document dstDoc = new Document();

        // Add a bookmark in the destination document where the nodes will be inserted.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.StartBookmark("InsertHere");
        builder.Write("Text before insertion.");
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
    }

    /// <summary>
    /// Inserts a collection of nodes into <paramref name="dstDoc"/> at the bookmark named <paramref name="bookmarkName"/>.
    /// </summary>
    /// <param name="dstDoc">The destination document.</param>
    /// <param name="srcDoc">The source document (used for importing).</param>
    /// <param name="nodes">The nodes to insert.</param>
    /// <param name="bookmarkName">The name of the bookmark that marks the insertion point.</param>
    static void InsertNodeCollectionAtBookmark(Document dstDoc, Document srcDoc, List<Node> nodes, string bookmarkName)
    {
        // Move the builder cursor to the start of the bookmark.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        bool found = builder.MoveToBookmark(bookmarkName);
        if (!found)
            throw new ArgumentException($"Bookmark '{bookmarkName}' not found in the destination document.");

        // Determine the node after which we will start inserting.
        Node insertionPoint = builder.CurrentParagraph ?? (Node)builder.CurrentNode;
        if (insertionPoint == null || (insertionPoint.NodeType != NodeType.Paragraph && insertionPoint.NodeType != NodeType.Table))
        {
            insertionPoint = dstDoc.Range.Bookmarks[bookmarkName].BookmarkStart;
        }

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
