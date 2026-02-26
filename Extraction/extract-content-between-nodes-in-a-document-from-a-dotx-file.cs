using System;
using System.Collections.Generic;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the DOTX template.
        Document template = new Document("Template.dotx");

        // Define the bookmark names that mark the start and end of the range to extract.
        const string startBookmarkName = "Start";
        const string endBookmarkName = "End";

        // Locate the start and end bookmark nodes.
        BookmarkStart startBookmark = template.Range.Bookmarks[startBookmarkName]?.BookmarkStart;
        BookmarkEnd endBookmark = template.Range.Bookmarks[endBookmarkName]?.BookmarkEnd;

        if (startBookmark == null || endBookmark == null)
        {
            Console.WriteLine("Start or End bookmark not found.");
            return;
        }

        // Create a new empty document that will hold the extracted content.
        Document extracted = new Document();
        // Remove the default section/paragraph that Aspose.Words creates.
        extracted.RemoveAllChildren();

        // Create a single section and body for the destination document.
        Section destSection = new Section(extracted);
        extracted.AppendChild(destSection);
        Body destBody = new Body(extracted);
        destSection.AppendChild(destBody);

        // Prepare a NodeImporter for efficient node copying.
        NodeImporter importer = new NodeImporter(template, extracted, ImportFormatMode.KeepSourceFormatting);

        // Collect all nodes that are between the start and end bookmarks (exclusive).
        List<Node> nodesBetween = new List<Node>();
        Node current = startBookmark.NextSibling;
        while (current != null && current != endBookmark)
        {
            nodesBetween.Add(current);
            current = current.NextSibling;
        }

        // Import the collected nodes into the new document.
        foreach (Node node in nodesBetween)
        {
            Node importedNode = importer.ImportNode(node, true);
            destBody.AppendChild(importedNode);
        }

        // Save the extracted content to a new DOCX file.
        extracted.Save("ExtractedContent.docx");
    }
}
