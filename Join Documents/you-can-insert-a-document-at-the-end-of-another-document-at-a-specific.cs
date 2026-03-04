using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class DocumentInsertionDemo
{
    static void Main()
    {
        // Paths to the documents – replace with actual file locations.
        const string destinationPath = @"C:\Docs\Destination.docx";
        const string sourcePath = @"C:\Docs\Source.docx";
        const string outputPathEnd = @"C:\Docs\Result_End.docx";
        const string outputPathBookmark = @"C:\Docs\Result_Bookmark.docx";
        const string outputPathParagraph = @"C:\Docs\Result_Paragraph.docx";
        const string outputPathInline = @"C:\Docs\Result_Inline.docx";

        // Load the destination and source documents.
        Document dstDoc = new Document(destinationPath);
        Document srcDoc = new Document(sourcePath);

        // -------------------------------------------------
        // 1. Insert the source document at the very end of the destination document.
        // -------------------------------------------------
        DocumentBuilder builderEnd = new DocumentBuilder(dstDoc);
        builderEnd.MoveToDocumentEnd();                                   // Move cursor to the end.
        builderEnd.InsertBreak(BreakType.PageBreak);                      // Optional page break before insertion.
        builderEnd.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        dstDoc.Save(outputPathEnd);                                        // Save the result.

        // -------------------------------------------------
        // 2. Insert the source document at a bookmark named "InsertHere".
        // -------------------------------------------------
        // Ensure the bookmark exists in the destination document.
        Document dstDocBookmark = new Document(destinationPath);
        DocumentBuilder builderBookmark = new DocumentBuilder(dstDocBookmark);
        builderBookmark.MoveToBookmark("InsertHere");                     // Position cursor at the bookmark.
        builderBookmark.InsertDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        dstDocBookmark.Save(outputPathBookmark);

        // -------------------------------------------------
        // 3. Insert the source document after a specific paragraph (e.g., paragraph index 2).
        // -------------------------------------------------
        Document dstDocParagraph = new Document(destinationPath);
        DocumentBuilder builderParagraph = new DocumentBuilder(dstDocParagraph);
        // Retrieve the third paragraph (zero‑based index).
        Paragraph targetParagraph = dstDocParagraph.FirstSection.Body.Paragraphs[2];
        builderParagraph.MoveTo(targetParagraph);                         // Move cursor to the paragraph.
        builderParagraph.InsertBreak(BreakType.PageBreak);                // Optional break.
        builderParagraph.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        dstDocParagraph.Save(outputPathParagraph);

        // -------------------------------------------------
        // 4. Insert the source document inline (the last paragraph break of the source is removed).
        // -------------------------------------------------
        Document dstDocInline = new Document(destinationPath);
        DocumentBuilder builderInline = new DocumentBuilder(dstDocInline);
        builderInline.MoveToDocumentEnd();                                 // Position at the end.
        // Use InsertDocumentInline to merge without an extra paragraph break.
        builderInline.InsertDocumentInline(srcDoc, ImportFormatMode.KeepSourceFormatting, new ImportFormatOptions());
        dstDocInline.Save(outputPathInline);
    }
}
