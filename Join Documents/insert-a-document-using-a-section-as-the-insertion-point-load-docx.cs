using System;
using Aspose.Words;

namespace AsposeWordsInsertExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the documents.
            string destinationPath = "Destination.docx";   // Existing document that will receive the insertion.
            string sourcePath = "Source.docx";           // Document whose content will be inserted.
            string resultPath = "Result.docx";           // Path where the merged document will be saved.

            // Load the destination and source documents (load rule).
            Document destinationDoc = new Document(destinationPath);
            Document sourceDoc = new Document(sourcePath);

            // Create a DocumentBuilder for the destination document (create rule).
            DocumentBuilder builder = new DocumentBuilder(destinationDoc);

            // Move the cursor to the beginning of the first section (insertion point).
            // You can change the index to target a different section.
            builder.MoveToSection(0);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the inserted content.
            builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting document (save rule).
            destinationDoc.Save(resultPath);
        }
    }
}
