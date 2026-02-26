using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsFormFieldInsert
{
    class Program
    {
        static void Main()
        {
            // Path to the template DOTX file (must be a Word template).
            const string templatePath = @"C:\Docs\Template.dotx";

            // Paths to the source documents that contain form fields.
            const string sourceDoc1Path = @"C:\Docs\FormFields1.docx";
            const string sourceDoc2Path = @"C:\Docs\FormFields2.docx";

            // Load the template document (DOTX) – this creates the base document.
            Document templateDoc = new Document(templatePath);

            // Create a DocumentBuilder for the template.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Move the cursor to the end of the template where we want to insert the content.
            builder.MoveToDocumentEnd();

            // Load the first source document that already contains form fields.
            Document srcDoc1 = new Document(sourceDoc1Path);

            // Insert the first document into the template.
            // Keep the original formatting of the source document.
            builder.InsertDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);

            // Insert a page break between the inserted documents (optional).
            builder.InsertBreak(BreakType.PageBreak);

            // Load the second source document that also contains form fields.
            Document srcDoc2 = new Document(sourceDoc2Path);

            // Insert the second document into the template.
            builder.InsertDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

            // Update all fields (including the inserted form fields) so that any calculated
            // results are refreshed before saving.
            templateDoc.UpdateFields();

            // Save the resulting document as a new DOTX template.
            const string outputPath = @"C:\Docs\ResultTemplate.dotx";
            templateDoc.Save(outputPath, SaveFormat.Dotx);
        }
    }
}
