using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractContentBetweenNodes
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Assume the document contains two bookmarks that mark the start and end of the region
        // you want to extract. The bookmarks are named "StartBookmark" and "EndBookmark".
        // Retrieve the bookmark objects.
        Bookmark startBookmark = doc.Range.Bookmarks["StartBookmark"];
        Bookmark endBookmark = doc.Range.Bookmarks["EndBookmark"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // Get the nodes that represent the start and end positions.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Collect the text of all nodes that lie between the start and end nodes (exclusive).
        StringBuilder extractedText = new StringBuilder();
        Node current = startNode;

        // Traverse forward until we reach the end node.
        while (current != null && current != endNode)
        {
            // Append the text of the current node.
            extractedText.Append(current.GetText());

            // Move to the next node in the document order.
            current = current.NextSibling;
        }

        // Optionally include the end node's text if needed:
        // extractedText.Append(endNode.GetText());

        // Output the extracted content.
        string result = extractedText.ToString().Trim();
        Console.WriteLine("Extracted Text:");
        Console.WriteLine(result);

        // Save the extracted text to a plain‑text file.
        File.WriteAllText("ExtractedContent.txt", result);
    }
}
