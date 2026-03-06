using System;
using Aspose.Words;

class InsertDocumentExample
{
    static void Main()
    {
        // Path to the source document that will be inserted.
        string sourcePath = "Source.docx";

        // Path where the resulting document will be saved.
        string resultPath = "Result.docx";

        // Load the source document from the file system.
        Document sourceDoc = new Document(sourcePath);

        // Create a new blank destination document.
        Document destinationDoc = new Document();

        // Use DocumentBuilder to work with the destination document.
        DocumentBuilder builder = new DocumentBuilder(destinationDoc);

        // Move the cursor to the end of the destination document.
        builder.MoveToDocumentEnd();

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to the specified file.
        destinationDoc.Save(resultPath);
    }
}
