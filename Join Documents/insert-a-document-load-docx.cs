using System;
using Aspose.Words;
using Aspose.Words.Loading;

class InsertDocumentExample
{
    static void Main()
    {
        // Paths to the source document to be inserted and the output document.
        string sourcePath = @"C:\Docs\Source.docx";
        string outputPath = @"C:\Docs\Result.docx";

        // Create a new blank document.
        Document destination = new Document();

        // Use DocumentBuilder to add initial content and position the cursor.
        DocumentBuilder builder = new DocumentBuilder(destination);
        builder.Writeln("Start of destination document.");

        // Load the source DOCX file.
        Document source = new Document(sourcePath);

        // Move the cursor to the end of the destination document.
        builder.MoveToDocumentEnd();

        // Insert the source document, keeping its original formatting.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        destination.Save(outputPath);
    }
}
