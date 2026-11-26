using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create the main document with a header that
        //    contains a bookmark named "InsertHere".
        // -------------------------------------------------
        var mainDoc = new Document();
        var mainSection = mainDoc.Sections[0];

        // Ensure the primary header exists.
        var header = mainSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
        {
            header = new HeaderFooter(mainDoc, HeaderFooterType.HeaderPrimary);
            mainSection.HeadersFooters.Add(header);
        }

        // Add a paragraph with an empty bookmark inside the header.
        var headerParagraph = new Paragraph(mainDoc);
        var bookmarkStart = new BookmarkStart(mainDoc, "InsertHere");
        var bookmarkEnd = new BookmarkEnd(mainDoc, "InsertHere");
        headerParagraph.AppendChild(bookmarkStart);
        headerParagraph.AppendChild(bookmarkEnd);
        header.AppendChild(headerParagraph);

        // Save the main document (sample input).
        const string mainPath = "MainDocument.docx";
        mainDoc.Save(mainPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the document whose content will be inserted.
        // -------------------------------------------------
        var insertDoc = new Document();
        var insertSection = insertDoc.Sections[0];

        // Add a couple of paragraphs as sample content.
        var para1 = new Paragraph(insertDoc);
        para1.AppendChild(new Run(insertDoc, "First inserted paragraph."));
        insertSection.Body.AppendChild(para1);

        var para2 = new Paragraph(insertDoc);
        para2.AppendChild(new Run(insertDoc, "Second inserted paragraph."));
        insertSection.Body.AppendChild(para2);

        // Save the insert document (sample input).
        const string insertPath = "InsertDocument.docx";
        insertDoc.Save(insertPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load both documents as if they were external files.
        // -------------------------------------------------
        var loadedMain = new Document(mainPath);
        var loadedInsert = new Document(insertPath);

        // -------------------------------------------------
        // 4. Locate the bookmark inside the header of the main document.
        // -------------------------------------------------
        var bookmark = loadedMain.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' was not found in the header.");

        // The insertion point should be after the paragraph that contains the bookmark.
        Node referenceNode = bookmark.BookmarkStart.ParentNode; // This is the Paragraph node.

        // -------------------------------------------------
        // 5. Import each node from the insert document and place it
        //    after the bookmark paragraph, preserving source formatting.
        // -------------------------------------------------
        foreach (Node node in loadedInsert.FirstSection.Body)
        {
            // Import the node into the main document.
            Node importedNode = loadedMain.ImportNode(node, true, ImportFormatMode.KeepSourceFormatting);

            // Insert the imported node after the reference paragraph.
            referenceNode.ParentNode.InsertAfter(importedNode, referenceNode);

            // Update the reference so the next node is inserted after this one.
            referenceNode = importedNode;
        }

        // -------------------------------------------------
        // 6. Save the resulting document, preserving header formatting.
        // -------------------------------------------------
        const string resultPath = "ResultDocument.docx";
        loadedMain.Save(resultPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 7. Validate that the output file exists.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("The resulting document was not saved correctly.", resultPath);
    }
}
