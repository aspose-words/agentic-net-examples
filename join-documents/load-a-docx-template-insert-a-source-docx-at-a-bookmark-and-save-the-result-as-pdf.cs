using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // Paths for the template, source, and output files.
        string templatePath = Path.Combine(workFolder, "Template.docx");
        string sourcePath = Path.Combine(workFolder, "Source.docx");
        string outputPath = Path.Combine(workFolder, "Result.pdf");

        // -----------------------------------------------------------------
        // Create the template document with a bookmark named "InsertHere".
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is the template document.");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("Bookmark location (will be replaced).");
        templateBuilder.EndBookmark("InsertHere");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the source document that will be inserted at the bookmark.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document inserted at the bookmark.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the template, locate the bookmark, and insert the source document.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Document loadedSource = new Document(sourcePath);

        // Ensure the bookmark exists.
        if (loadedTemplate.Range.Bookmarks["InsertHere"] == null)
        {
            throw new InvalidOperationException("Bookmark 'InsertHere' not found in the template.");
        }

        // Move the builder to the bookmark and insert the source document.
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedTemplate);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(loadedSource, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged result as PDF.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Failed to create the PDF output file.");
        }

        // Indicate success (no interactive prompts required).
        Console.WriteLine($"PDF created successfully at: {outputPath}");
    }
}
