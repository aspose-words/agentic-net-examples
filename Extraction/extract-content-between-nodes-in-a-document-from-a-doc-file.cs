using System;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the source DOC file.
        Document sourceDoc = new Document("Input.doc");

        // Assume the document contains two bookmarks that define the range to extract:
        // "Start" marks the beginning and "End" marks the end.
        Bookmark startBookmark = sourceDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = sourceDoc.Range.Bookmarks["End"];

        // Get the actual nodes that delimit the range.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Create a new blank document that will hold the extracted fragment.
        Document fragmentDoc = new Document();

        // Use NodeImporter to copy nodes from the source document to the fragment
        // while preserving the original formatting.
        NodeImporter importer = new NodeImporter(sourceDoc, fragmentDoc, ImportFormatMode.KeepSourceFormatting);

        // Walk through the sibling nodes that lie between the start and end bookmarks.
        // Exclude the bookmark nodes themselves.
        Node current = startNode.NextSibling;
        while (current != null && current != endNode)
        {
            // Import the node into the fragment document.
            Node importedNode = importer.ImportNode(current, true);
            // Append the imported node to the body of the fragment document.
            fragmentDoc.FirstSection.Body.AppendChild(importedNode);
            // Move to the next sibling.
            current = current.NextSibling;
        }

        // The extracted text is now available via the fragment's Range.Text property.
        string extractedText = fragmentDoc.Range.Text.Trim();

        Console.WriteLine("Extracted text between bookmarks:");
        Console.WriteLine(extractedText);

        // Optionally, save the extracted fragment as a separate DOCX file.
        fragmentDoc.Save("ExtractedFragment.docx");
    }
}
