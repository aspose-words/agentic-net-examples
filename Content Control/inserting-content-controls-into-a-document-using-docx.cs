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

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Insert an inline plain‑text content control (SDT).
        // ------------------------------------------------------------
        StructuredDocumentTag plainTextTag = new StructuredDocumentTag(
            doc,                     // The document the tag belongs to.
            SdtType.PlainText,      // Type of the content control.
            MarkupLevel.Inline);    // Inline placement.

        // Optional: set friendly title and tag identifier.
        plainTextTag.Title = "PlainTextControl";
        plainTextTag.Tag = "PlainTextTag";

        // Insert the content control into the document.
        builder.InsertNode(plainTextTag);

        // Add placeholder text that will appear inside the control.
        builder.Writeln("Enter text here...");

        // Add a paragraph break after the control.
        builder.Writeln();

        // ------------------------------------------------------------
        // Insert a block‑level rich‑text content control (SDT).
        // ------------------------------------------------------------
        StructuredDocumentTag richTextTag = new StructuredDocumentTag(
            doc,
            SdtType.RichText,
            MarkupLevel.Block); // Block‑level placement.

        richTextTag.Title = "RichTextControl";
        richTextTag.Tag = "RichTextTag";

        // Insert the block‑level control.
        builder.InsertNode(richTextTag);

        // Add some default content inside the rich‑text control.
        builder.Writeln("Rich text content goes here.");

        // ------------------------------------------------------------
        // Save the document as a DOCX file.
        // ------------------------------------------------------------
        doc.Save("ContentControls.docx", SaveFormat.Docx);
    }
}
