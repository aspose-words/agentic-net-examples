using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the primary document (the one that will receive the other document).
        Document mainDoc = new Document("Source1.docx");

        // Load the document that we want to insert.
        Document docToInsert = new Document("Source2.docx");

        // Create a DocumentBuilder for the primary document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Position the builder at the very end of the primary document.
        builder.MoveToDocumentEnd();

        // Optional: insert a page break so the inserted content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the second document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted document.
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to a new file.
        mainDoc.Save("Combined.docx");
    }
}
