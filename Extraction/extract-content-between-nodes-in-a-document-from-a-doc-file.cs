using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOC file.
        Document doc = new Document("Input.doc");

        // Identify the start and end nodes. In this example we use two bookmarks named "Start" and "End".
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // Nodes that mark the boundaries.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Collect text that lies between the two nodes (excluding the bookmark markers themselves).
        StringBuilder extracted = new StringBuilder();

        // Begin with the node immediately after the start bookmark.
        Node cur = startNode.NextSibling;

        while (cur != null && cur != endNode)
        {
            extracted.Append(cur.GetText());
            cur = cur.NextSibling;
        }

        // Output the extracted content.
        Console.WriteLine("Extracted text:");
        Console.WriteLine(extracted.ToString());
    }
}
