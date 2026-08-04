using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with sample content.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source Paragraph 1");
        srcBuilder.Writeln("Source Paragraph 2");
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln("Source Paragraph 3");

        // Create a destination document with a custom bookmark.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination before bookmark.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Placeholder paragraph that will be replaced.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination after bookmark.");

        // Extract the nodes (paragraphs and tables) from the source document.
        NodeCollection sourceNodes = sourceDoc.FirstSection.Body.GetChildNodes(NodeType.Any, true);
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Locate the bookmark in the destination document.
        Bookmark bookmark = destDoc.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' was not found in the destination document.");

        // The bookmark is inside a paragraph. We'll use that paragraph as the insertion point.
        Paragraph bookmarkParagraph = bookmark.BookmarkStart.ParentNode as Paragraph;
        if (bookmarkParagraph == null)
            throw new InvalidOperationException("Bookmark is not located inside a paragraph.");

        // Remove the placeholder paragraph that follows the bookmark.
        Node placeholder = bookmarkParagraph.NextSibling;
        if (placeholder != null && placeholder.NodeType == NodeType.Paragraph)
            placeholder.Remove();

        // Insert imported nodes after the bookmark paragraph.
        CompositeNode body = destDoc.FirstSection.Body;
        Node insertionReference = bookmarkParagraph;

        foreach (Node node in sourceNodes)
        {
            if (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table)
            {
                Node importedNode = importer.ImportNode(node, true);
                body.InsertAfter(importedNode, insertionReference);
                insertionReference = importedNode;
            }
        }

        // Save the resulting document.
        string outputPath = "Result.docx";
        destDoc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The result document was not created.");

        Console.WriteLine($"Document created successfully at '{Path.GetFullPath(outputPath)}'.");
    }
}
