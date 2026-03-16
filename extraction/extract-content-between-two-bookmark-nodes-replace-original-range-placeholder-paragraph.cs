using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class BookmarkExtractor
{
    static void Main()
    {
        // Load the source document.
        Document srcDoc = new Document("Input.docx");

        // Name of the bookmark that encloses the content to extract.
        const string bookmarkName = "ContentBookmark";

        // Retrieve the bookmark.
        Bookmark bookmark = srcDoc.Range.Bookmarks[bookmarkName];
        if (bookmark == null)
            throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

        // Nodes that lie between the start and end of the bookmark.
        Node startNode = bookmark.BookmarkStart;
        Node endNode   = bookmark.BookmarkEnd;

        // Collect the nodes to extract (excluding the bookmark markers themselves).
        List<Node> nodesToExtract = new List<Node>();
        for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
            nodesToExtract.Add(cur);

        // -----------------------------------------------------------------
        // 1. Create a new document that will contain the extracted content.
        // -----------------------------------------------------------------
        Document extractedDoc = new Document();
        CompositeNode dstStory = extractedDoc.FirstSection.Body;

        // Use NodeImporter for efficient import of nodes from the source document.
        NodeImporter importer = new NodeImporter(srcDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in nodesToExtract)
        {
            Node imported = importer.ImportNode(node, true);
            dstStory.AppendChild(imported);
        }

        // Save the extracted fragment (optional, for verification).
        extractedDoc.Save("ExtractedContent.docx");

        // ---------------------------------------------------------------
        // 2. Remove the original nodes from the source document.
        // ---------------------------------------------------------------
        foreach (Node node in nodesToExtract)
            node.Remove();

        // ---------------------------------------------------------------
        // 3. Insert a placeholder paragraph where the original content was.
        // ---------------------------------------------------------------
        Paragraph placeholder = new Paragraph(srcDoc);
        Run placeholderRun = new Run(srcDoc, "[Placeholder]");
        placeholder.AppendChild(placeholderRun);

        // Insert the placeholder after the bookmark start marker.
        // InsertAfter is defined on CompositeNode, so we need the parent as CompositeNode.
        CompositeNode parent = (CompositeNode)startNode.ParentNode;
        parent.InsertAfter(placeholder, startNode);

        // Finally, remove the bookmark markers themselves.
        startNode.Remove();
        endNode.Remove();

        // Save the modified source document.
        srcDoc.Save("OutputWithPlaceholder.docx");
    }
}
