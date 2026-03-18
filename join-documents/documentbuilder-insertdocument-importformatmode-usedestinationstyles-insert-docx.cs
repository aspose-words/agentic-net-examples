using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace InsertDocAndSaveHtml
{
    class Program
    {
        static void Main()
        {
            // Paths to the files – replace with actual locations.
            string templatePath = @"C:\Docs\StyledTemplate.docx";   // The styled template.
            string sourcePath   = @"C:\Docs\ContentToInsert.docx"; // The document to be inserted.
            string outputPath   = @"C:\Docs\Result.html";          // The final HTML file.

            // Load the template (styled) document.
            Document templateDoc = new Document(templatePath);

            // Load the source document that will be inserted.
            Document sourceDoc = new Document(sourcePath);

            // Create a DocumentBuilder for the template document.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Move the cursor to the end of the template (or any other position you need).
            builder.MoveToDocumentEnd();

            // Insert the source document using the UseDestinationStyles mode.
            // This keeps the template's styles and copies any new styles from the source.
            builder.InsertDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

            // Save the combined document as HTML.
            templateDoc.Save(outputPath, SaveFormat.Html);
        }
    }
}
