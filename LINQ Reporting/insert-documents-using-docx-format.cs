using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank destination document.
        Document destination = new Document();

        // Attach a DocumentBuilder to the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the cursor to the end of the document and insert a page break
        // to separate any existing content from the inserted document.
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Load the source DOCX document that will be inserted.
        Document source = new Document("Source.docx");

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to a new DOCX file.
        destination.Save("CombinedResult.docx");
    }
}
