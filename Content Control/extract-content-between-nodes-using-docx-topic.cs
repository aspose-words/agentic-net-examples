using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the source DOCX document.
        Document srcDoc = new Document("Input.docx");

        // Define the start and end bookmarks that enclose the desired content.
        // Ensure that the document contains bookmarks named "Start" and "End".
        Bookmark startBookmark = srcDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = srcDoc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // The actual nodes that mark the boundaries.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Create a new empty document that will hold the extracted content.
        Document destDoc = new Document();

        // Ensure the destination document has at least one section and body.
        destDoc.EnsureMinimum();

        // Use NodeImporter for efficient import of nodes from source to destination.
        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Start with the node immediately after the start bookmark.
        Node curNode = startNode.NextSibling;

        // Iterate through the source nodes until we reach the end bookmark.
        while (curNode != null && curNode != endNode)
        {
            // Keep a reference to the next sibling before we import the current node,
            // because importing creates a clone that is independent of the source tree.
            Node nextNode = curNode.NextSibling;

            // Import the node (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(curNode, true);

            // Append the imported node to the body of the destination document.
            destDoc.FirstSection.Body.AppendChild(importedNode);

            // Move to the next node in the source document.
            curNode = nextNode;
        }

        // Save the extracted content to a new DOCX file.
        destDoc.Save("ExtractedContent.docx");
    }
}
