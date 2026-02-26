using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractBetweenBookmarks
{
    static void Main()
    {
        // Load the DOCM file (lifecycle rule: load)
        Document doc = new Document("InputDocument.docm");

        // Retrieve the start and end bookmarks that define the region to extract
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark = doc.Range.Bookmarks["End"];

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Required bookmarks not found.");
            return;
        }

        // Get the underlying bookmark nodes
        Node startNode = startBookmark.BookmarkStart;
        Node endNode = endBookmark.BookmarkEnd;

        // Collect text from all nodes that appear after the start node and before the end node
        StringBuilder extractedText = new StringBuilder();
        Node current = startNode;

        // Traverse the document using pre‑order traversal until we reach the end bookmark node
        while (current != null && current != endNode)
        {
            // Append the text of the current node (including its children)
            extractedText.Append(current.GetText());
            current = current.NextPreOrder(doc);
        }

        // Create a new document to hold the extracted content (lifecycle rule: create)
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Write the extracted text into the new document
        builder.Writeln(extractedText.ToString().Trim());

        // Save the result (lifecycle rule: save)
        resultDoc.Save("ExtractedContent.docx");

        Console.WriteLine("Extraction complete. Saved to ExtractedContent.docx");
    }
}
