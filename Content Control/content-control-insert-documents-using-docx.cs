using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertDocumentIntoContentControl
{
    static void Main()
    {
        // Path to the folder that contains the source documents.
        string dataDir = @"C:\Docs\";

        // Load the main document that will receive the inserted content.
        Document mainDoc = new Document(dataDir + "Main.docx");

        // Load the document that will be inserted.
        Document subDoc = new Document(dataDir + "Sub.docx");

        // Create a DocumentBuilder for the main document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Insert a block‑level rich‑text content control (StructuredDocumentTag) at the current cursor position.
        StructuredDocumentTag sdt = new StructuredDocumentTag(mainDoc, SdtType.RichText, MarkupLevel.Block);
        builder.InsertNode(sdt);

        // Move the builder's cursor inside the newly created content control.
        builder.MoveTo(sdt);

        // Insert the contents of the sub‑document into the content control.
        // Keep the original formatting of the source document.
        builder.InsertDocument(subDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        mainDoc.Save(dataDir + "Result.docx");
    }
}
