using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the RTF document.
        Document doc = new Document("input.rtf");

        // Identify the start and end nodes.
        // In this example we use bookmarks named "Start" and "End".
        // Adjust the bookmark names as needed for your document.
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // The actual nodes that mark the boundaries.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Collect text between the two nodes (excluding the bookmark markers themselves).
        StringBuilder extractedText = new StringBuilder();
        Node current = startNode.NextSibling;

        while (current != null && current != endNode)
        {
            // Append the text of the current node and its children.
            extractedText.Append(current.GetText());
            current = current.NextSibling;
        }

        // Output the extracted content.
        Console.WriteLine(extractedText.ToString());
    }
}
