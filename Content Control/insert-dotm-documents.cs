using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDotmExample
{
    static void Main()
    {
        // Create a new blank document (destination).
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Optionally move the cursor to the end and add a page break before insertion.
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Load the DOTM (macro‑enabled template) that we want to insert.
        Document srcDoc = new Document("Template.dotm");

        // Insert the DOTM content into the destination document.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }
}
