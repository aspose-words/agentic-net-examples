using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag plainTag = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        // Add placeholder text inside the content control.
        builder.Write("Enter your name here");
        // Exit the content control and start a new paragraph.
        builder.Writeln();

        // Insert a rich‑text content control.
        StructuredDocumentTag richTag = builder.InsertStructuredDocumentTag(SdtType.RichText);
        builder.Write("Rich text content control");
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("ContentControls.docx");
    }
}
