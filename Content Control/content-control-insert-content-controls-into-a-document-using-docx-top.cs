using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Plain Text Content Control (block level) -----
        StructuredDocumentTag plainTextControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        plainTextControl.Title = "PlainTextCC";
        plainTextControl.Tag = "PlainTextTag";
        // Insert the content control into the document.
        builder.InsertNode(plainTextControl);
        // Move the cursor inside the newly created content control.
        builder.MoveTo(plainTextControl);
        // Add text that will be inside the plain‑text content control.
        builder.Writeln("This text is inside a plain text content control.");

        // ----- Rich Text Content Control (block level) -----
        StructuredDocumentTag richTextControl = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richTextControl.Title = "RichTextCC";
        richTextControl.Tag = "RichTextTag";
        builder.InsertNode(richTextControl);
        builder.MoveTo(richTextControl);
        builder.Writeln("This text is inside a rich text content control.");

        // Save the document in DOCX format.
        doc.Save("ContentControls.docx", SaveFormat.Docx);
    }
}
