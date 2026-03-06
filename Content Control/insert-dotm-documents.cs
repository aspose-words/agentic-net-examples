using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDotmExample
{
    static void Main()
    {
        // Load the DOTM (macro‑enabled template) that we want to insert.
        Document dotmTemplate = new Document("Template.dotm");

        // Create a new blank document that will receive the inserted content.
        Document destDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // Position the builder at the end of the document (or any other location as needed).
        builder.MoveToDocumentEnd();

        // Optional: insert a page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the DOTM document into the destination document.
        // KeepSourceFormatting preserves the original styles and formatting of the template.
        builder.InsertDocument(dotmTemplate, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        destDoc.Save("Result.docx");
    }
}
