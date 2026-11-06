using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the main document with a header that contains a bookmark.
        Document mainDoc = new Document();
        Section mainSection = mainDoc.Sections[0];

        // Add a header to the first section.
        HeaderFooter header = new HeaderFooter(mainDoc, HeaderFooterType.HeaderPrimary);
        mainSection.HeadersFooters.Add(header);

        // Build a paragraph in the header with a bookmark.
        Paragraph headerPara = new Paragraph(mainDoc);
        Run runBefore = new Run(mainDoc, "Header before bookmark. ");
        headerPara.AppendChild(runBefore);

        // Insert bookmark.
        BookmarkStart bookmarkStart = new BookmarkStart(mainDoc, "InsertHere");
        BookmarkEnd bookmarkEnd = new BookmarkEnd(mainDoc, "InsertHere");
        headerPara.AppendChild(bookmarkStart);
        headerPara.AppendChild(new Run(mainDoc, " ")); // placeholder for inserted content.
        headerPara.AppendChild(bookmarkEnd);

        Run runAfter = new Run(mainDoc, " Header after bookmark.");
        headerPara.AppendChild(runAfter);
        header.AppendChild(headerPara);

        // Save the main document (optional, for debugging).
        const string mainPath = "Main.docx";
        mainDoc.Save(mainPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        Paragraph insertPara = new Paragraph(insertDoc);
        insertPara.AppendChild(new Run(insertDoc, "This is the inserted content from the DOCX."));
        insertDoc.FirstSection.Body.AppendChild(insertPara);

        // Save the insert document (optional, for debugging).
        const string insertPath = "Insert.docx";
        insertDoc.Save(insertPath, SaveFormat.Docx);

        // Load the documents (simulating a real scenario where they are separate files).
        Document srcDoc = new Document(mainPath);
        Document docToInsert = new Document(insertPath);

        // Locate the bookmark inside the header.
        Bookmark bookmark = srcDoc.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found.");

        // The bookmark is inside a paragraph; we will insert after the bookmark start.
        Node insertAfterNode = bookmark.BookmarkStart;

        // Import each node from the document to insert into the main document.
        // We will import the body of the first section (all its child nodes).
        foreach (Node srcNode in docToInsert.FirstSection.Body)
        {
            Node importedNode = srcDoc.ImportNode(srcNode, true, ImportFormatMode.KeepSourceFormatting);
            // Insert the imported node after the current node.
            srcDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].AppendChild(importedNode);
        }

        // Save the resulting document preserving header formatting.
        const string resultPath = "Result.docx";
        srcDoc.Save(resultPath, SaveFormat.Docx);

        // Validate that the output file exists.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("The result document was not created.", resultPath);
    }
}
