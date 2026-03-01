using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a descriptive paragraph.
        builder.Writeln("Below is a rich‑text content control (Structured Document Tag):");
        builder.Writeln(); // Add an empty line for spacing.

        // Insert a rich‑text content control.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.RichText);

        // Configure the content control's properties.
        sdt.Title = "CustomerInfo";          // Title shown in the UI.
        sdt.Tag = "CustomerInfoTag";         // Tag used for identification.
        sdt.PlaceholderName = "Enter customer details here";

        // Insert default text inside the content control.
        builder.Writeln("John Doe, 123 Main St.");

        // Ensure the cursor is at the end of the document before saving.
        builder.MoveToDocumentEnd();

        // Save the document in DOCX format.
        doc.Save("ContentControlExample.docx");
    }
}
