using System;
using Aspose.Words;

namespace InsertDocumentDynamically
{
    class Program
    {
        static void Main()
        {
            // Path to the main template document that contains a bookmark placeholder.
            // The bookmark name is "InsertHere". It can be created in Word beforehand.
            string templatePath = @"C:\Docs\Template.docx";

            // Path to the external document that we want to insert.
            string externalDocPath = @"C:\Docs\External.docx";

            // Load the template document.
            Document templateDoc = new Document(templatePath);               // Document(string)

            // Load the external document.
            Document externalDoc = new Document(externalDocPath);            // Document(string)

            // Create a DocumentBuilder for the template document.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);      // DocumentBuilder(Document)

            // Move the cursor to the bookmark where the external document will be inserted.
            builder.MoveToBookmark("InsertHere");                           // MoveToBookmark(string)

            // Insert the external document at the bookmark position.
            // KeepSourceFormatting preserves the original formatting of the inserted document.
            builder.InsertDocument(externalDoc, ImportFormatMode.KeepSourceFormatting); // InsertDocument(Document, ImportFormatMode)

            // Save the resulting document.
            string outputPath = @"C:\Docs\Result.docx";
            templateDoc.Save(outputPath);                                   // Save(string)
        }
    }
}
