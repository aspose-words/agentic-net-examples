using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the RTF document using RtfLoadOptions.
        string rtfPath = "input.rtf";
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document(rtfPath, loadOptions);

        // The document is expected to contain two bookmarks named "Start" and "End".
        // These bookmarks define the region whose content we want to extract.
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];

        // Retrieve the underlying nodes for the bookmark start and end.
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Collect the text that lies between the two bookmark nodes.
        StringBuilder extracted = new StringBuilder();

        // Walk the document in pre‑order traversal, beginning after the start node,
        // and stop when the end node is reached.
        Node current = startNode;
        while ((current = current.NextPreOrder(doc)) != null && current != endNode)
        {
            // Append the textual representation of each node.
            extracted.Append(current.GetText());
        }

        string result = extracted.ToString().Trim();

        Console.WriteLine("Extracted text between bookmarks:");
        Console.WriteLine(result);
    }
}
