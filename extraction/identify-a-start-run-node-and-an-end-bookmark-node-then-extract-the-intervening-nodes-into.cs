using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add some initial text.
        builder.Writeln("Paragraph before the start run.");

        // Insert the start run that will be used as the beginning marker.
        builder.Write("StartRun");
        // Keep a reference to this run node.
        Run startRun = (Run)sourceDoc.GetChild(NodeType.Run,
            sourceDoc.GetChildNodes(NodeType.Run, true).Count - 1, true);

        // Add more content that will be extracted.
        builder.Writeln(); // empty paragraph
        builder.Writeln("First paragraph inside the range.");
        builder.Writeln("Second paragraph inside the range.");

        // Insert a bookmark that will serve as the end marker.
        builder.StartBookmark("EndBookmark");
        builder.Writeln("Content inside the bookmark (will not be extracted).");
        builder.EndBookmark("EndBookmark");

        // Add some content after the bookmark.
        builder.Writeln("Paragraph after the bookmark.");

        // Locate the end bookmark node (BookmarkEnd) and its containing paragraph.
        Bookmark endBookmark = sourceDoc.Range.Bookmarks["EndBookmark"];
        BookmarkEnd endBookmarkNode = endBookmark.BookmarkEnd;
        Paragraph endParagraph = endBookmarkNode.ParentNode as Paragraph;

        // Locate the paragraph that contains the start run.
        Paragraph startParagraph = startRun.ParentNode as Paragraph;

        // Collect all block-level nodes that lie between the start paragraph and the end paragraph (exclusive).
        List<Node> nodesToExtract = new List<Node>();
        Node currentNode = startParagraph.NextSibling;
        while (currentNode != null && currentNode != endParagraph)
        {
            nodesToExtract.Add(currentNode);
            currentNode = currentNode.NextSibling;
        }

        // Validate that we have extracted nodes.
        if (nodesToExtract.Count == 0)
            throw new InvalidOperationException("No nodes were found between the start run and the end bookmark.");

        // Create a new destination document with a clean structure.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // remove the default empty section/paragraph

        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Prepare a NodeImporter for cloning nodes from source to destination.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each collected node into the destination document.
        foreach (Node node in nodesToExtract)
        {
            // All collected nodes are block-level (Paragraph, Table, etc.).
            Node importedNode = importer.ImportNode(node, true);
            destBody.AppendChild(importedNode);
        }

        // Save the extracted content to a file.
        const string outputPath = "Extracted.docx";
        destDoc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The extracted document was not saved correctly.");

        Console.WriteLine($"Extraction completed. Output saved to '{outputPath}'.");
    }
}
