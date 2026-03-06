using System;
using Aspose.Words;

namespace InsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination documents.
            string dataDir = @"C:\Docs\";
            string destinationPath = dataDir + "Destination.docx";
            string sourcePath = dataDir + "Source.docx";
            string resultPath = dataDir + "Result.docx";

            // Load the destination document (the document into which we will insert another document).
            Document destinationDoc = new Document(destinationPath);

            // Load the source document (the document that will be inserted).
            Document sourceDoc = new Document(sourcePath);

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destinationDoc);

            // Move the cursor to the end of the destination document.
            builder.MoveToDocumentEnd();

            // Optionally insert a page break before the inserted content.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the source document.
            builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            destinationDoc.Save(resultPath);
        }
    }
}
