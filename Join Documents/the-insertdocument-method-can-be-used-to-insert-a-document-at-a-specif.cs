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

            // Load the destination document (create rule).
            Document destinationDoc = new Document(destinationPath);

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destinationDoc);

            // Move the cursor to the end of the document where we want to insert the source.
            builder.MoveToDocumentEnd();

            // Optionally insert a page break before the inserted content.
            builder.InsertBreak(BreakType.PageBreak);

            // Load the source document that will be inserted (load rule).
            Document sourceDoc = new Document(sourcePath);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the inserted content.
            builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document (save rule).
            destinationDoc.Save(resultPath);
        }
    }
}
