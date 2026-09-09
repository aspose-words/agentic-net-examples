using System;
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

        builder.Writeln("Paragraph before start.");
        // Paragraph that contains the start run.
        builder.Writeln(); // create empty paragraph
        builder.Write("StartRun"); // this run will be the start marker
        builder.Writeln(); // end of the paragraph

        builder.Writeln("Paragraph 1 between markers.");
        builder.Writeln("Paragraph 2 between markers.");

        // Insert a table to demonstrate mixed content extraction.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell B1");
        builder.EndRow();
        builder.EndTable();

        // End bookmark marker.
        builder.StartBookmark("EndRange");
        builder.Writeln("Paragraph after end bookmark.");
        builder.EndBookmark("EndRange");

        // Save the source document (optional, for inspection).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Locate the start Run node with the exact text "StartRun".
        Run startRun = null;
        foreach (Run run in sourceDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "StartRun")
            {
                startRun = run;
                break;
            }
        }

        if (startRun == null)
            throw new InvalidOperationException("Start run not found.");

        // Locate the end bookmark node (BookmarkEnd) named "EndRange".
        Bookmark endBookmark = sourceDoc.Range.Bookmarks["EndRange"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");

        BookmarkEnd endBookmarkNode = endBookmark.BookmarkEnd;
        if (endBookmarkNode == null)
            throw new InvalidOperationException("End bookmark node not found.");

        // Determine the block-level nodes that bound the extraction range.
        Paragraph startParagraph = startRun.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmarkNode.ParentNode as Paragraph;

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to determine paragraph boundaries.");

        // Prepare the destination document.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren();

        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);

        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Importer to handle node import between documents.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Traverse sibling nodes between the start and end paragraphs (exclusive).
        Node currentNode = startParagraph.NextSibling;
        while (currentNode != null && currentNode != endParagraph)
        {
            // Clone and import the node into the destination document.
            Node importedNode = importer.ImportNode(currentNode, true);
            destBody.AppendChild(importedNode);
            currentNode = currentNode.NextSibling;
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        destDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed: output file was not created.");

        // Optional: indicate success (no console interaction required).
        // The program will exit normally.
    }
}
