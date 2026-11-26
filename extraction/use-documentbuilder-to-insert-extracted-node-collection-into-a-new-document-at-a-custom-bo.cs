using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ---------- Create a source document with sample paragraphs ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source Paragraph 1");
        srcBuilder.Writeln("Source Paragraph 2");
        srcBuilder.Writeln("Source Paragraph 3");
        sourceDoc.Save("source.docx");

        // ---------- Load the source document and get its paragraphs ----------
        Document loadedSource = new Document("source.docx");
        NodeCollection sourceParagraphs = loadedSource.GetChildNodes(NodeType.Paragraph, true);
        if (sourceParagraphs == null || sourceParagraphs.Count == 0)
            throw new InvalidOperationException("No paragraphs were extracted from the source document.");

        // ---------- Create a destination document with a custom bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination start.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination end.");
        destDoc.Save("dest.docx");

        // ---------- Load the destination document ----------
        Document loadedDest = new Document("dest.docx");

        // ---------- Locate the bookmark and its containing paragraph ----------
        Bookmark bookmark = loadedDest.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found in destination document.");

        // The BookmarkStart node resides inside a Paragraph.
        Paragraph insertionParagraph = bookmark.BookmarkStart.ParentNode as Paragraph;
        if (insertionParagraph == null)
            throw new InvalidOperationException("Unable to locate the paragraph that contains the bookmark.");

        // ---------- Prepare an importer to bring nodes from source into destination ----------
        NodeImporter importer = new NodeImporter(loadedSource, loadedDest, ImportFormatMode.KeepSourceFormatting);

        // ---------- Import and insert each paragraph after the bookmark ----------
        // Keep a reference to the last inserted node so subsequent inserts follow it.
        Node lastInsertedNode = insertionParagraph;
        foreach (Paragraph para in sourceParagraphs)
        {
            // Import the paragraph into the destination document.
            Node importedNode = importer.ImportNode(para, true);
            if (importedNode == null)
                throw new InvalidOperationException("Failed to import a paragraph node.");

            // Insert the imported paragraph after the previously inserted node.
            lastInsertedNode.ParentNode.InsertAfter(importedNode, lastInsertedNode);
            lastInsertedNode = importedNode;
        }

        // ---------- Save the resulting document ----------
        loadedDest.Save("result.docx");

        // ---------- Verify that the result file was created ----------
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not created.");
    }
}
