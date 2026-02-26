using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Rich Text content control (StructuredDocumentTag).
        StructuredDocumentTag richTag = builder.InsertStructuredDocumentTag(SdtType.RichText);
        richTag.Title = "CustomerName";
        richTag.Tag = "CustomerNameTag";

        // Add placeholder text inside the rich text content control.
        builder.Writeln("John Doe");

        // Move the cursor out of the current content control.
        builder.MoveToDocumentEnd();

        // Insert a Plain Text content control.
        StructuredDocumentTag plainTag = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        plainTag.Title = "OrderNumber";
        plainTag.Tag = "OrderNumberTag";

        // Add placeholder text inside the plain text content control.
        builder.Writeln("12345");

        // Save the document to a DOCX file.
        doc.Save("ContentControl.docx");
    }
}
