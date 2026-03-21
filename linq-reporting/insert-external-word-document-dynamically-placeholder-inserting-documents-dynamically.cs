using System;
using System.IO;
using Aspose.Words;

namespace InsertDocumentDynamically
{
    class Program
    {
        static void Main()
        {
            // Create the template document with a bookmark placeholder.
            Document templateDoc = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
            templateBuilder.Writeln("This is the main document.");
            templateBuilder.StartBookmark("InsertHere");
            templateBuilder.Writeln("[Placeholder for external document]");
            templateBuilder.EndBookmark("InsertHere");
            templateBuilder.Writeln("End of the main document.");

            // Create the external document that will be inserted.
            Document externalDoc = new Document();
            DocumentBuilder externalBuilder = new DocumentBuilder(externalDoc);
            externalBuilder.Writeln("This is the content of the external document.");
            externalBuilder.Writeln("It will be inserted at the bookmark.");

            // Insert the external document at the bookmark position.
            templateBuilder.MoveToBookmark("InsertHere");
            templateBuilder.InsertDocument(externalDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
            templateDoc.Save(outputPath);

            Console.WriteLine($"Result document saved to: {outputPath}");
        }
    }
}
