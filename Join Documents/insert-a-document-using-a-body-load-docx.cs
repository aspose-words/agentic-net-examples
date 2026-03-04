using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Insert the source document into the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.MoveToDocumentEnd(); // Position cursor at the end.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("Combined.docx");
    }
}
