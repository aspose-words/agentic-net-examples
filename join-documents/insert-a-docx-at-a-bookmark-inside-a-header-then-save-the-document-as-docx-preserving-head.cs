using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the documents that will be created and the final result.
        string destPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string resultPath = "Result.docx";

        // -------------------------------------------------
        // 1. Create the destination document with a header that contains a bookmark.
        // -------------------------------------------------
        Document destDoc = new Document();
        Section destSection = destDoc.Sections[0];

        HeaderFooter header = new HeaderFooter(destDoc, HeaderFooterType.HeaderPrimary);
        destSection.HeadersFooters.Add(header);

        Paragraph headerParagraph = new Paragraph(destDoc);
        headerParagraph.AppendChild(new Run(destDoc, "Header before bookmark "));
        // Bookmark where the source document will be inserted.
        BookmarkStart bookmarkStart = new BookmarkStart(destDoc, "InsertHere");
        BookmarkEnd bookmarkEnd = new BookmarkEnd(destDoc, "InsertHere");
        headerParagraph.AppendChild(bookmarkStart);
        headerParagraph.AppendChild(bookmarkEnd);
        headerParagraph.AppendChild(new Run(destDoc, " Header after bookmark"));
        header.AppendChild(headerParagraph);

        // Save the destination document (optional, just to have a file on disk).
        destDoc.Save(destPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the source document that will be inserted.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is inserted content from the source document.");
        srcBuilder.Writeln("Second paragraph of inserted content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Locate the bookmark inside the header.
        // -------------------------------------------------
        BookmarkStart? bmStart = null;
        foreach (BookmarkStart bm in header.GetChildNodes(NodeType.BookmarkStart, true))
        {
            if (bm.Name == "InsertHere")
            {
                bmStart = bm;
                break;
            }
        }

        if (bmStart == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' was not found in the header.");

        // The bookmark resides inside a paragraph.
        Paragraph bmParagraph = (Paragraph)bmStart.ParentNode;
        // The parent of that paragraph is the HeaderFooter node.
        HeaderFooter headerParent = (HeaderFooter)bmParagraph.ParentNode;

        // -------------------------------------------------
        // 4. Insert the source document's body nodes after the bookmark paragraph.
        // -------------------------------------------------
        Node? lastInserted = bmParagraph;
        foreach (Node srcNode in sourceDoc.FirstSection.Body)
        {
            // Import each node while keeping its original formatting.
            Node importedNode = destDoc.ImportNode(srcNode, true, ImportFormatMode.KeepSourceFormatting);
            headerParent.InsertAfter(importedNode, lastInserted);
            lastInserted = importedNode;
        }

        // -------------------------------------------------
        // 5. Save the final document preserving header formatting.
        // -------------------------------------------------
        destDoc.Save(resultPath, SaveFormat.Docx);

        // Verify that the result file was created.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("The resulting document was not saved.", resultPath);
    }
}
