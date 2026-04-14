using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the template, source and result files.
        string templatePath = Path.Combine(dataDir, "Template.docx");
        string sourcePath   = Path.Combine(dataDir, "Source.docx");
        string resultPath   = Path.Combine(dataDir, "Result.pdf");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a bookmark.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        templateBuilder.Writeln("This is the template document.");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("[Placeholder for inserted document]");
        templateBuilder.EndBookmark("InsertHere");

        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a source document that will be inserted at the bookmark.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        sourceBuilder.Writeln("This is the inserted source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the template, move to the bookmark and insert the source.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(templatePath);
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

        // Position the cursor at the bookmark.
        mainBuilder.MoveToBookmark("InsertHere");

        // Load the source document and insert it.
        Document srcToInsert = new Document(sourcePath);
        mainBuilder.InsertDocument(srcToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the merged document as PDF.
        // -----------------------------------------------------------------
        mainDoc.Save(resultPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(templatePath);
        // File.Delete(sourcePath);
    }
}
