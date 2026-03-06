using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertContentControls
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (StructuredDocumentTag).
        // The method returns the created tag, which we can use if further configuration is required.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Write some text inside the content control.
        builder.Write("This text is inside a plain‑text content control.");

        // Close the content control by moving the cursor out of it.
        // InsertStructuredDocumentTag creates a container; after writing, the cursor is positioned inside it.
        // Adding a paragraph break moves the cursor outside the tag.
        builder.Writeln();

        // Save the document to a DOCX file.
        doc.Save("ContentControls.docx");
    }
}
