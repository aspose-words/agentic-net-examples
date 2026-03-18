using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocAtBookmarkAndSavePdf
{
    static void Main()
    {
        // Paths to the template, the document to insert, and the output PDF.
        string templatePath = @"C:\Docs\Template.docx";
        string sourcePath   = @"C:\Docs\Source.docx";
        string outputPath   = @"C:\Docs\Result.pdf";

        // Load the template document (create/load rule).
        Document templateDoc = new Document(templatePath);

        // Load the source document that will be inserted.
        Document sourceDoc = new Document(sourcePath);

        // Create a DocumentBuilder for the template.
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Move the cursor to the bookmark named "InsertHere".
        // The bookmark must exist in the template; otherwise an exception is thrown.
        builder.MoveToBookmark("InsertHere");

        // Insert the source document at the bookmark position.
        // Keep the source formatting while inserting.
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF (save rule).
        templateDoc.Save(outputPath, SaveFormat.Pdf);
    }
}
