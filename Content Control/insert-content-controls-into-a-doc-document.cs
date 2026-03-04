using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertContentControls
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control.
        // SdtType.PlainText creates a simple text placeholder.
        builder.InsertStructuredDocumentTag(SdtType.PlainText);
        builder.Writeln("This is inside a plain‑text content control.");

        // Insert a rich‑text content control.
        // SdtType.RichText allows formatted content inside the control.
        builder.InsertStructuredDocumentTag(SdtType.RichText);
        builder.Writeln("This is inside a rich‑text content control.");

        // Insert a picture content control.
        // SdtType.Picture creates a placeholder for an image.
        builder.InsertStructuredDocumentTag(SdtType.Picture);
        builder.Writeln("Picture content control placeholder.");

        // Save the document to a DOCX file.
        doc.Save("ContentControls.docx");
    }
}
