using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocumentExample
{
    static void Main()
    {
        // Path to the folder that contains the documents.
        string docsPath = @"C:\Docs\";

        // Create a new blank document (uses Document() constructor).
        Document destination = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Load the source DOCX document from file (uses Document(string) constructor).
        Document source = new Document(docsPath + "Source.docx");

        // Move the cursor to the end of the destination document.
        builder.MoveToDocumentEnd();

        // Insert the source document at the current cursor position,
        // keeping the source formatting (uses InsertDocument method).
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to a new file (uses Document.Save(string) method).
        destination.Save(docsPath + "Combined.docx");
    }
}
