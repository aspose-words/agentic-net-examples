using System;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the source PDF file into an Aspose.Words Document.
        Document sourceDoc = new Document("input.pdf");

        // Assume the PDF contains two bookmarks named "Start" and "End"
        // that mark the beginning and the end of the region we want to extract.
        Bookmark startBookmark = sourceDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = sourceDoc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // Create a new empty document that will hold the extracted content.
        Document extractedDoc = new Document();

        // NodeImporter efficiently copies nodes from the source document to the target document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Get the actual nodes that delimit the range.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Walk through the sibling chain from the start node up to (and including) the end node.
        Node curNode = startNode;
        while (curNode != null)
        {
            // Import the current node into the target document.
            Node importedNode = importer.ImportNode(curNode, true);
            // Append the imported node to the body of the first section of the target document.
            extractedDoc.FirstSection.Body.AppendChild(importedNode);

            // Stop after processing the end node.
            if (curNode == endNode)
                break;

            curNode = curNode.NextSibling;
        }

        // Save the extracted fragment to a new file (DOCX format in this example).
        extractedDoc.Save("extracted.docx");
        Console.WriteLine("Extraction completed successfully.");
    }
}
