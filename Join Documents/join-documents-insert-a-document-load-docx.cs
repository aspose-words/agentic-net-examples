using System;
using Aspose.Words;

namespace AsposeWordsJoinExample
{
    class Program
    {
        static void Main()
        {
            // Path to the main document that will receive the inserted content.
            string mainDocPath = "MainDocument.docx";

            // Path to the document that will be inserted.
            string docToInsertPath = "DocumentToInsert.docx";

            // Path where the resulting joined document will be saved.
            string outputPath = "JoinedDocument.docx";

            // Load the main document (creates a Document object from an existing DOCX file).
            Document mainDoc = new Document(mainDocPath);

            // Load the document that we want to insert.
            Document docToInsert = new Document(docToInsertPath);

            // Create a DocumentBuilder for the main document.
            DocumentBuilder builder = new DocumentBuilder(mainDoc);

            // Move the cursor to the end of the main document (optional, can be any position).
            builder.MoveToDocumentEnd();

            // Insert a page break to separate the original content from the inserted content.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the second document at the current cursor position.
            // ImportFormatMode.KeepSourceFormatting preserves the formatting of the inserted document.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document to the specified output file.
            mainDoc.Save(outputPath);
        }
    }
}
