using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsInsertAndExport
{
    class Program
    {
        static void Main()
        {
            // Create a temporary folder for the output.
            string tempFolder = Path.GetTempPath();
            string outputPath = Path.Combine(tempFolder, "Result.html");

            // -----------------------------------------------------------------
            // Create the template document with two bookmarks: Header and Footer.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

            templateBuilder.Writeln("Template start");

            // Header bookmark placeholder.
            templateBuilder.StartBookmark("Header");
            templateBuilder.Writeln("[Header placeholder]");
            templateBuilder.EndBookmark("Header");

            templateBuilder.Writeln("Middle content");

            // Footer bookmark placeholder.
            templateBuilder.StartBookmark("Footer");
            templateBuilder.Writeln("[Footer placeholder]");
            templateBuilder.EndBookmark("Footer");

            templateBuilder.Writeln("Template end");

            // --------------------------------------------------------------
            // Create the document whose content will be inserted at each bookmark.
            // --------------------------------------------------------------
            Document insertDoc = new Document();
            DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
            insertBuilder.Writeln("Inserted line 1");
            insertBuilder.Writeln("Inserted line 2");
            insertBuilder.Writeln("Inserted line 3");

            // --------------------------------------------------------------
            // Insert the content at the "Header" bookmark.
            // --------------------------------------------------------------
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.MoveToBookmark("Header");
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // --------------------------------------------------------------
            // Insert the same content at the "Footer" bookmark.
            // --------------------------------------------------------------
            builder.MoveToBookmark("Footer");
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // --------------------------------------------------------------
            // Save the resulting document as HTML.
            // --------------------------------------------------------------
            templateDoc.Save(outputPath, SaveFormat.Html);

            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
