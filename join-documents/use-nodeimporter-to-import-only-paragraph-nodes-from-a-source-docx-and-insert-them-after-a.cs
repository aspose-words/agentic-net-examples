using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create destination document with a bookmark.
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document Start");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark location");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination Document End");

        // Save destination document (optional, just to have a file).
        const string destPath = "Destination.docx";
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create source document containing several paragraphs.
        var srcDoc = new Document();
        var srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source Paragraph 1");
        srcBuilder.Writeln("Source Paragraph 2");
        srcBuilder.Writeln("Source Paragraph 3");

        // Save source document (optional, just to have a file).
        const string srcPath = "Source.docx";
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // Load documents (in case they were saved to disk).
        var destination = new Document(destPath);
        var source = new Document(srcPath);

        // Locate the bookmark in the destination document.
        if (!destination.Range.Bookmarks.Any(b => b.Name == "InsertHere"))
            throw new InvalidOperationException("Bookmark 'InsertHere' not found in destination document.");

        var bookmark = destination.Range.Bookmarks["InsertHere"];
        var bookmarkParagraph = bookmark.BookmarkStart.ParentNode as Paragraph;
        if (bookmarkParagraph == null)
            throw new InvalidOperationException("Bookmark does not have a parent paragraph.");

        // Prepare a NodeImporter to import nodes from source to destination.
        var importer = new NodeImporter(source, destination, ImportFormatMode.KeepSourceFormatting);

        // Insert each paragraph from the source document after the bookmark.
        Node insertAfterNode = bookmarkParagraph;
        foreach (Paragraph srcParagraph in source.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
        {
            // Import the paragraph node (including its children).
            Node importedNode = importer.ImportNode(srcParagraph, true);
            // Insert the imported paragraph after the previous node.
            insertAfterNode.ParentNode?.InsertAfter(importedNode, insertAfterNode);
            // Update the reference so the next paragraph is inserted after this one.
            insertAfterNode = importedNode;
        }

        // Save the resulting document.
        const string resultPath = "Result.docx";
        destination.Save(resultPath, SaveFormat.Docx);

        // Validate that the result file was created.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("Result document was not created.", resultPath);

        // Optional validation: ensure the expected number of paragraphs were inserted.
        int sourceParagraphCount = source.GetChildNodes(NodeType.Paragraph, true).Count;
        int resultParagraphCount = destination.GetChildNodes(NodeType.Paragraph, true).Count;
        // Destination originally had 3 paragraphs (start, bookmark location, end).
        if (resultParagraphCount < sourceParagraphCount + 3)
            throw new InvalidOperationException("Not all source paragraphs were inserted into the destination document.");
    }
}
