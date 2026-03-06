using System;
using Aspose.Words;

class InsertDocumentExample
{
    static void Main()
    {
        // Path to the folder that contains the source and destination files.
        string docsPath = @"C:\Docs\";

        // Load the source document that we want to insert.
        Document srcDoc = new Document(docsPath + "Source.docx");   // load rule

        // Create a new blank destination document.
        Document dstDoc = new Document();                         // create rule

        // Attach a DocumentBuilder to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);    // create rule

        // Insert the source document at the current cursor position (beginning of the document).
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting); // insert rule

        // Save the combined document.
        dstDoc.Save(docsPath + "Result.docx");                   // save rule
    }
}
