using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create destination document with a bookmark.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document start.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark location.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination document end.");

        // Save destination for reference (optional).
        string destPath = "Destination.docx";
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create source document containing several paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source paragraph 1.");
        srcBuilder.Writeln("Source paragraph 2.");
        srcBuilder.Writeln("Source paragraph 3.");

        // Save source for reference (optional).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the source document (simulating a separate load step).
        Document srcToImport = new Document(sourcePath);

        // Prepare a NodeImporter to import nodes from source to destination.
        NodeImporter importer = new NodeImporter(srcToImport, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Locate the bookmark in the destination document.
        Bookmark bookmark = destDoc.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found in destination document.");

        // The paragraph that contains the bookmark start.
        Paragraph bookmarkParagraph = (Paragraph)bookmark.BookmarkStart.ParentNode;

        // Insert imported paragraphs after the bookmark paragraph.
        Node insertAfterNode = bookmarkParagraph;
        foreach (Paragraph srcParagraph in srcToImport.GetChildNodes(NodeType.Paragraph, true))
        {
            // Import the paragraph node.
            Node importedNode = importer.ImportNode(srcParagraph, true);
            // Insert after the previously inserted node.
            Node insertedNode = insertAfterNode.ParentNode.InsertAfter(importedNode, insertAfterNode);
            // Update reference for next insertion.
            insertAfterNode = insertedNode;
        }

        // Save the merged result.
        string resultPath = "Result.docx";
        destDoc.Save(resultPath, SaveFormat.Docx);

        // Validation: ensure the result file exists.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("The merged document was not saved correctly.", resultPath);

        // Validation: ensure the result contains the expected number of paragraphs.
        int expectedParagraphs = destDoc.GetChildNodes(NodeType.Paragraph, true).Count;
        Document resultDoc = new Document(resultPath);
        int actualParagraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true).Count;
        if (actualParagraphs < expectedParagraphs)
            throw new InvalidOperationException("The merged document does not contain the expected paragraphs.");

        // Program completed successfully.
    }
}
