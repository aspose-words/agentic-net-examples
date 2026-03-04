using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOCX file (lifecycle: load)
        Document doc = new Document("Input.docx");

        // The document must contain two bookmarks named "Start" and "End"
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Bookmarks 'Start' and/or 'End' not found in the document.");
            return;
        }

        // Nodes that mark the boundaries
        Node startNode = startBookmark.BookmarkStart;
        Node endNode   = endBookmark.BookmarkEnd;

        // Extract text that lies between the two bookmarks
        string extracted = GetTextBetween(startNode, endNode, doc);
        Console.WriteLine("Extracted text between bookmarks:");
        Console.WriteLine(extracted);
    }

    // Returns the concatenated text of all nodes that are strictly between startNode and endNode.
    static string GetTextBetween(Node startNode, Node endNode, Document doc)
    {
        StringBuilder sb = new StringBuilder();

        // Move to the first node after the start bookmark using the overload that requires the root node.
        Node current = startNode.NextPreOrder(doc);

        // Walk the document tree until we reach the end bookmark (exclusive).
        while (current != null && !current.Equals(endNode))
        {
            sb.Append(current.GetText());
            current = current.NextPreOrder(doc);
        }

        return sb.ToString();
    }
}
