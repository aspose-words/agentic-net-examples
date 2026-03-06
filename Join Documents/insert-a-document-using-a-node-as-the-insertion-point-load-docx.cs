using System;
using Aspose.Words;

namespace AsposeWordsInsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the documents.
            string destinationPath = @"C:\Docs\Destination.docx";
            string sourcePath = @"C:\Docs\Source.docx";
            string outputPath = @"C:\Docs\Result.docx";

            // Load the destination document (the document that will receive the insertion).
            Document destinationDoc = new Document(destinationPath);

            // Load the source document (the document to be inserted).
            Document sourceDoc = new Document(sourcePath);

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destinationDoc);

            // Move the cursor to the insertion point.
            // In this example we use a bookmark named "InsertHere" that must exist in the destination document.
            builder.MoveToBookmark("InsertHere");

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the inserted content.
            builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the modified document.
            destinationDoc.Save(outputPath);
        }
    }
}
