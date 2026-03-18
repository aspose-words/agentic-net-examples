using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsInsertAndExport
{
    class Program
    {
        static void Main()
        {
            // Paths to the source (template) document and the document that will be inserted.
            // The template must contain two bookmarks named "Header" and "Footer".
            const string templatePath = @"C:\Docs\Template.docx";
            const string insertPath   = @"C:\Docs\InsertContent.docx";
            const string outputPath   = @"C:\Docs\Result.html";

            // Load the template document.
            Document templateDoc = new Document(templatePath);

            // Load the document whose whole content will be inserted at each bookmark.
            Document insertDoc = new Document(insertPath);

            // Create a DocumentBuilder for the template document.
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert the content at the "Header" bookmark.
            builder.MoveToBookmark("Header");
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Insert the same content at the "Footer" bookmark.
            builder.MoveToBookmark("Footer");
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting document as HTML.
            // Using the overload that specifies the format directly.
            templateDoc.Save(outputPath, SaveFormat.Html);
        }
    }
}
