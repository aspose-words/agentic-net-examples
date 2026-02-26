using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document mainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Add some introductory text.
        builder.Writeln("Document before the inserted content:");

        // Insert a RichText content control (StructuredDocumentTag) into the document.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.RichText);

        // Position the builder's cursor inside the newly created content control.
        builder.MoveTo(sdt);

        // Load the external DOCX file that we want to embed.
        Document docToInsert = new Document("Insert.docx");

        // Insert the external document at the current cursor position,
        // preserving the source formatting of the inserted content.
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Move the cursor to the end of the main document and add concluding text.
        builder.MoveToDocumentEnd();
        builder.Writeln("\nDocument after the inserted content.");

        // Save the resulting document.
        mainDoc.Save("Result.docx");
    }
}
