using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Plain‑text content control (inline) ----------
        // Create the content control.
        StructuredDocumentTag plainTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        // Insert it into the document.
        builder.InsertNode(plainTag);
        // Move the builder inside the newly created tag.
        builder.MoveTo(plainTag);
        // Add the text that will be inside the control.
        builder.Writeln("This text is inside a plain‑text content control.");

        // ---------- Rich‑text content control (block) ----------
        StructuredDocumentTag richTag = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        builder.InsertNode(richTag);
        builder.MoveTo(richTag);
        builder.Writeln("This paragraph is inside a rich‑text content control.");

        // Save the document.
        doc.Save("ContentControls.docx");
    }
}
