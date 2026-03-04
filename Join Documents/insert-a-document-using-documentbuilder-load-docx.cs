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
        Document srcDoc = new Document(sourcePath); // uses Document(string) constructor

        // Create a new blank destination document.
        Document dstDoc = new Document(); // uses Document() constructor

        // Attach a DocumentBuilder to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc); // uses DocumentBuilder(Document) constructor

        // Insert the source document at the current cursor position (beginning of the document).
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting); // uses InsertDocument method

        // Save the combined document to the specified file.
        dstDoc.Save(resultPath); // uses Save(string) method
    }
}
