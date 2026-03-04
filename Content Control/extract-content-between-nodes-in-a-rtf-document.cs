using System;
using System.IO;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the RTF document (lifecycle rule: load)
        Document doc = new Document("InputDocument.rtf");

        // Assume the start and end positions are marked by bookmarks named "Start" and "End".
        // Retrieve the bookmark nodes.
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // Get the nodes that delimit the range.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Build the text that lies between the two nodes.
        // Walk the document tree from the node after the start bookmark up to (but not including) the end bookmark.
        Node current = startNode.NextSibling;
        var extractedText = new System.Text.StringBuilder();

        while (current != null && current != endNode)
        {
            // Append the text of each node (including its children).
            extractedText.Append(current.GetText());
            current = current.NextSibling;
        }

        // Output the extracted content.
        string result = extractedText.ToString();
        Console.WriteLine("Extracted text between bookmarks:");
        Console.WriteLine(result);

        // Optionally, save the extracted text to a plain‑text file (lifecycle rule: save).
        File.WriteAllText("ExtractedContent.txt", result);
    }
}
