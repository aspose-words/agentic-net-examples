using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractBetweenBookmarks
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Define the names of the start and end bookmarks that enclose the desired content.
        const string startBookmarkName = "Start";
        const string endBookmarkName = "End";

        // Extract the text that lies between the two bookmarks.
        string extractedText = GetTextBetweenBookmarks(doc, startBookmarkName, endBookmarkName);

        // Output the extracted text to the console.
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(extractedText);

        // Optionally, save the extracted text to a plain‑text file.
        File.WriteAllText("Extracted.txt", extractedText);
    }

    /// <summary>
    /// Returns the concatenated text of all nodes that appear after the start bookmark
    /// and before the end bookmark in the given document.
    /// </summary>
    static string GetTextBetweenBookmarks(Document doc, string startBookmark, string endBookmark)
    {
        // Retrieve the bookmark objects from the document.
        Bookmark start = doc.Range.Bookmarks[startBookmark];
        Bookmark end = doc.Range.Bookmarks[endBookmark];

        // Validate that both bookmarks exist.
        if (start == null || end == null)
            throw new ArgumentException("One or both bookmarks were not found in the document.");

        // The BookmarkStart and BookmarkEnd nodes are the actual nodes in the document tree.
        Node startNode = start.BookmarkStart;
        Node endNode = end.BookmarkEnd;

        // Use a StringBuilder to accumulate the text.
        StringBuilder sb = new StringBuilder();

        // Traverse the document tree in pre‑order, starting with the node that follows the start bookmark.
        for (Node cur = startNode.NextPreOrder(startNode.Document); cur != null && cur != endNode; cur = cur.NextPreOrder(startNode.Document))
        {
            // Skip the end bookmark node itself.
            if (cur == endNode)
                break;

            // Append the visible text of the current node.
            sb.Append(cur.GetText());
        }

        // Return the collected text, trimmed of leading/trailing whitespace.
        return sb.ToString().Trim();
    }
}
