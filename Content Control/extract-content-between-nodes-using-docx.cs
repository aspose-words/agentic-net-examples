using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the source DOCX document (lifecycle rule: load)
        Document srcDoc = new Document("Source.docx");

        // Assume the document contains two bookmarks named "Start" and "End"
        // that mark the region we want to extract.
        Bookmark startBookmark = srcDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = srcDoc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // The actual nodes that delimit the range.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Create a new blank document to hold the extracted content (lifecycle rule: create)
        Document extractedDoc = new Document();

        // The new document must have at least one section, body and paragraph.
        // Use the constructor that creates a minimal document (already satisfied).

        // Prepare a NodeImporter for efficient node copying (feature rule: use existing API)
        NodeImporter importer = new NodeImporter(srcDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert the imported nodes after the first paragraph of the new document.
        Paragraph insertionPoint = extractedDoc.FirstSection.Body.FirstParagraph;

        // Iterate from the node after the start bookmark up to (and including) the node before the end bookmark.
        Node curNode = startNode.NextPreOrder(srcDoc);
        while (curNode != null && !curNode.Equals(endNode))
        {
            // Import the node into the destination document.
            Node importedNode = importer.ImportNode(curNode, true);
            // Insert after the current insertion point and update the insertion point.
            insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
            insertionPoint = (Paragraph)importedNode; // advance insertion point
            // Move to the next node in the source document.
            curNode = curNode.NextPreOrder(srcDoc);
        }

        // Save the extracted content to a new DOCX file (lifecycle rule: save)
        extractedDoc.Save("ExtractedContent.docx", SaveFormat.Docx);
    }
}
