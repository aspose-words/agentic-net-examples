using System;
using System.Linq;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Create a source document programmatically.
        Document srcDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(srcDoc);

        // Add some content before the start run.
        builder.Writeln("Paragraph before start.");

        // Insert the start run containing the marker text "START".
        builder.Writeln("START");

        // Insert a bookmark named "MyBookmark" with some content.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Content inside bookmark.");
        builder.EndBookmark("MyBookmark");

        // Add some content after the bookmark.
        builder.Writeln("Paragraph after bookmark.");

        // Identify the start Run node (the first run that contains the text "START").
        Run startRun = srcDoc.GetChildNodes(NodeType.Run, true)
                             .Cast<Run>()
                             .FirstOrDefault(r => r.Text.Contains("START"));
        if (startRun == null)
            throw new InvalidOperationException("Start run not found.");

        // Identify the end Bookmark by name and obtain its BookmarkEnd node.
        Bookmark endBookmark = srcDoc.Range.Bookmarks["MyBookmark"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");
        Node endNode = endBookmark.BookmarkEnd;

        // Create a new empty document that will hold the extracted nodes.
        Document destDoc = new Document();

        // Prepare an importer to copy nodes from the source to the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Walk through the sibling chain starting after the start run up to (but not including) the end bookmark.
        Node curNode = startRun.NextSibling;
        while (curNode != null && curNode != endNode)
        {
            // Import the current node into the destination document.
            Node importedNode = importer.ImportNode(curNode, true);
            // Append the imported node to the body of the destination document.
            destDoc.FirstSection.Body.AppendChild(importedNode);
            // Move to the next sibling.
            curNode = curNode.NextSibling;
        }

        // Save the extracted fragment as a separate document.
        destDoc.Save("ExtractedFragment.docx");
        Console.WriteLine("Extraction completed. File saved as ExtractedFragment.docx");
    }
}
