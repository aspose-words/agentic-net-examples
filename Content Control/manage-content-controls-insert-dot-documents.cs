using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Markup;

class InsertDotIntoContentControl
{
    static void Main()
    {
        // Create a new blank document.
        Document dstDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Insert a plain‑text Structured Document Tag (content control) where the DOT will be placed.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        // Optionally give the content control a title for easier identification.
        sdt.Title = "InsertedTemplate";

        // Load the source DOT (Word template) document.
        Document srcDoc = new Document("Template.dot");

        // Move the builder's cursor inside the newly created content control.
        builder.MoveTo(sdt);

        // Insert the entire source document into the content control.
        // KeepSourceFormatting preserves the formatting defined in the DOT.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        dstDoc.Save("Result.docx");
    }
}
