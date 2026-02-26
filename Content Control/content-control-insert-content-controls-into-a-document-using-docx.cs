using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (StructuredDocumentTag) at the current cursor position.
        // The builder is automatically positioned inside the newly created content control.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Optionally set properties of the content control.
        sdt.Title = "SamplePlainTextControl";
        sdt.Tag = "PlainTextTag";

        // Write some placeholder text inside the content control.
        builder.Writeln("This text is inside a plain‑text content control.");

        // Move the cursor out of the content control to continue normal document editing.
        builder.MoveToDocumentEnd();

        // Insert a rich‑text content control as another example.
        StructuredDocumentTag richSdt = builder.InsertStructuredDocumentTag(SdtType.RichText);
        richSdt.Title = "SampleRichTextControl";
        richSdt.Tag = "RichTextTag";

        // Add formatted content inside the rich‑text control.
        builder.Font.Bold = true;
        builder.Writeln("Bold text inside a rich‑text content control.");
        builder.Font.Bold = false;
        builder.Writeln("Normal text inside the same control.");

        // Save the document in DOCX format.
        doc.Save("ContentControl.docx");
    }
}
