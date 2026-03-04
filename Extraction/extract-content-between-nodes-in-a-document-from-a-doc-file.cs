using System;
using System.Text;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // Load the document (uses the Document(string) constructor – a provided load rule).
        Document sourceDoc = new Document(sourcePath);

        // Define the names of the bookmarks that mark the start and end of the region to extract.
        // These bookmarks must exist in the source document.
        const string startBookmarkName = "Start";
        const string endBookmarkName   = "End";

        // Locate the bookmark start nodes using XPath (SelectSingleNode is a provided method).
        Node startNode = sourceDoc.SelectSingleNode($"//BookmarkStart[@Name='{startBookmarkName}']");
        Node endNode   = sourceDoc.SelectSingleNode($"//BookmarkStart[@Name='{endBookmarkName}']");

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("One or both bookmarks were not found.");
            return;
        }

        // Accumulate the text that lies between the two bookmark nodes.
        StringBuilder extractedText = new StringBuilder();

        // Start with the node immediately after the start bookmark.
        // NextPreOrder requires the root node of the traversal (the document).
        Node curNode = startNode.NextPreOrder(sourceDoc);

        // Walk the document tree until we reach the end bookmark.
        while (curNode != null && curNode != endNode)
        {
            // GetText returns the text of the node and its children.
            extractedText.Append(curNode.GetText());
            curNode = curNode.NextPreOrder(sourceDoc);
        }

        // Create a new blank document (uses the Document() constructor – a provided create rule).
        Document resultDoc = new Document();

        // Use DocumentBuilder to insert the extracted text into the new document.
        DocumentBuilder builder = new DocumentBuilder(resultDoc);
        builder.Writeln(extractedText.ToString().Trim());

        // Save the result (uses the Save(string) method – a provided save rule).
        string resultPath = @"C:\Docs\ExtractedContent.docx";
        resultDoc.Save(resultPath);

        Console.WriteLine($"Extracted content saved to: {resultPath}");
    }
}
