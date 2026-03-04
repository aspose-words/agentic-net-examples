using System;
using System.Collections.Generic;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOTX template.
        Document srcDoc = new Document("Template.dotx");

        // Assume the document contains two bookmarks named "Start" and "End"
        // that define the region we want to extract.
        Bookmark startBookmark = srcDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = srcDoc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Start or End bookmark not found.");
            return;
        }

        // Get the actual nodes that mark the boundaries.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Collect all nodes that lie between the start and end nodes (exclusive).
        List<Node> nodesBetween = new List<Node>();
        Node cur = startNode.NextPreOrder(srcDoc);
        while (cur != null && cur != endNode)
        {
            nodesBetween.Add(cur);
            cur = cur.NextPreOrder(srcDoc);
        }

        // Create a new blank document to hold the extracted content.
        Document destDoc = new Document();
        // Ensure the destination document has at least one section and a body.
        if (destDoc.FirstSection == null)
            destDoc.AppendChild(new Section(destDoc));
        if (destDoc.FirstSection.Body == null)
            destDoc.FirstSection.AppendChild(new Body(destDoc));

        // Use NodeImporter for efficient import of nodes from the source to the destination.
        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Append each imported node to the body of the destination document.
        foreach (Node node in nodesBetween)
        {
            Node importedNode = importer.ImportNode(node, true);
            destDoc.FirstSection.Body.AppendChild(importedNode);
        }

        // Save the extracted content as a separate DOCX file.
        destDoc.Save("ExtractedContent.docx");
    }
}
