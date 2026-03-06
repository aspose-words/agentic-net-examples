using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original document.
            Document originalDoc = new Document("Original.docx");

            // Clone the original document (deep copy).
            Document clonedDoc = originalDoc.Clone();

            // Load a document that will be inserted into the cloned document.
            Document docToInsert = new Document("Insert.docx");

            // Use DocumentBuilder to position the cursor at the end of the cloned document.
            DocumentBuilder builder = new DocumentBuilder(clonedDoc);
            builder.MoveToDocumentEnd();

            // Optionally insert a page break before the inserted content.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the document at the current cursor position.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Load a document that will be appended to the cloned document.
            Document docToAppend = new Document("Append.docx");

            // Append the document to the end of the cloned document.
            clonedDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting document after cloning, inserting and appending.
            clonedDoc.Save("Result_Cloned_Inserted_Appended.docx");

            // Split the cloned document: extract pages 1 through 2 into a new document.
            Document splitDoc = clonedDoc.ExtractPages(1, 2);

            // Save the split document.
            splitDoc.Save("Result_Split_Pages_1_2.docx");
        }
    }
}
