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

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a plain‑text content control (StructuredDocumentTag).
        // SdtType.PlainText creates a simple text control.
        // MarkupLevel.Inline places the control inside a paragraph.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        sdt.Title = "MyPlainTextControl";   // Friendly name shown in the UI.
        sdt.Tag = "PlainTextTag";           // Custom tag identifier.

        // Insert the content control at the current builder position.
        builder.InsertNode(sdt);

        // Add placeholder text inside the content control.
        sdt.AppendChild(new Run(doc, "Enter text here..."));

        // Save the document as a DOCX file using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        doc.Save("ContentControl.docx", saveOptions);
    }
}
