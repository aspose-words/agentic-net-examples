using System;
using System.IO;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOCM file from disk.
        Document doc = new Document("input.docm");

        // ------------------------------------------------------------
        // Example 1: Extract text that is enclosed in a bookmark.
        // ------------------------------------------------------------
        // The bookmark name "Extract" should surround the desired content.
        // Bookmark.Text returns the concatenated text inside the bookmark.
        string extractedText = string.Empty;
        if (doc.Range.Bookmarks["Extract"] != null)
        {
            extractedText = doc.Range.Bookmarks["Extract"].Text;
        }

        // ------------------------------------------------------------
        // Example 2: Extract text between two arbitrary nodes using XPath.
        // ------------------------------------------------------------
        // Uncomment the following block if you prefer node‑based extraction
        // instead of using a bookmark.
        /*
        // Select the start and end nodes (e.g., first and third paragraphs).
        Node startNode = doc.SelectSingleNode("//w:p[1]");
        Node endNode   = doc.SelectSingleNode("//w:p[3]");

        // Collect text from all nodes that appear after startNode and before endNode.
        if (startNode != null && endNode != null)
        {
            Node curNode = startNode.NextPreOrder(doc);
            while (curNode != null && curNode != endNode)
            {
                extractedText += curNode.GetText();
                curNode = curNode.NextPreOrder(doc);
            }
        }
        */

        // Save the extracted content to a plain‑text file.
        File.WriteAllText("extracted.txt", extractedText);
    }
}
