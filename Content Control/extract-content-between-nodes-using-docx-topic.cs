using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsExtraction
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document srcDoc = new Document("Input.docx");

            // Define the start and end bookmarks that delimit the content to extract.
            // Ensure that the document contains bookmarks named "Start" and "End".
            Bookmark startBookmark = srcDoc.Range.Bookmarks["Start"];
            Bookmark endBookmark = srcDoc.Range.Bookmarks["End"];

            if (startBookmark == null || endBookmark == null)
                throw new InvalidOperationException("Required bookmarks not found.");

            // The actual nodes that mark the boundaries.
            Node startNode = startBookmark.BookmarkStart;
            Node endNode = endBookmark.BookmarkEnd;

            // Create a new empty document that will hold the extracted fragment.
            Document extractedDoc = new Document();
            // Remove the default empty section/paragraph created by the constructor.
            extractedDoc.RemoveAllChildren();

            // Add a new section and body to the empty document.
            Section section = new Section(extractedDoc);
            extractedDoc.AppendChild(section);
            Body body = new Body(extractedDoc);
            section.AppendChild(body);

            // Use NodeImporter for efficient import of nodes from the source document.
            NodeImporter importer = new NodeImporter(srcDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

            // Iterate over sibling nodes starting after the start bookmark until the end bookmark.
            Node currentNode = startNode.NextSibling;
            while (currentNode != null && currentNode != endNode)
            {
                // Import the node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(currentNode, true);
                body.AppendChild(importedNode);

                // Move to the next sibling.
                currentNode = currentNode.NextSibling;
            }

            // Save the extracted content to a new DOCX file.
            extractedDoc.Save("ExtractedContent.docx");
        }
    }
}
