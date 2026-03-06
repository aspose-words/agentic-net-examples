using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Load the .dot template (Word template files can be opened as regular documents).
        Document dotDoc = new Document("Template.dot");

        // Move the builder cursor to the end of the destination document.
        builder.MoveToDocumentEnd();

        // Insert a page break before the template content (optional).
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the .dot document into the destination document,
        // preserving the source formatting.
        builder.InsertDocument(dotDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        dstDoc.Save("Result.docx");
    }
}
