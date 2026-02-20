using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting nodes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (Structured Document Tag).
        // The third argument is the markup level (Inline or Block). Use Inline for a plain‑text Sdt.
        StructuredDocumentTag plainTextControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        builder.InsertNode(plainTextControl);
        // Move the builder inside the control to add its content.
        builder.MoveTo(plainTextControl);
        builder.Writeln("This text is inside a plain‑text content control.");

        // Insert a rich‑text content control.
        builder.Writeln(); // start a new paragraph.
        // Use Inline markup level for the rich‑text Sdt as well.
        StructuredDocumentTag richTextControl = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline);
        builder.InsertNode(richTextControl);
        builder.MoveTo(richTextControl);
        builder.Writeln("This text is inside a rich‑text content control.");

        // Save the document as DOCX using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        doc.Save("ContentControls.docx", saveOptions);
    }
}
