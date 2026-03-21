using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocAtBookmarkAndSavePdf
{
    static void Main()
    {
        // Create a template document with a bookmark named "InsertHere".
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("Template start");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("[Bookmark location]");
        templateBuilder.EndBookmark("InsertHere");
        templateBuilder.Writeln("Template end");

        // Create a source document that will be inserted.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the inserted document content.");

        // Move the cursor to the bookmark in the template.
        DocumentBuilder insertBuilder = new DocumentBuilder(templateDoc);
        insertBuilder.MoveToBookmark("InsertHere");

        // Insert the source document at the bookmark position, keeping its formatting.
        insertBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Define the output PDF path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.pdf");

        // Save the combined document as PDF.
        templateDoc.Save(outputPath, SaveFormat.Pdf);

        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
