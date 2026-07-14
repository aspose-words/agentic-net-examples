using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.pdf");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a bookmark.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        templateBuilder.Writeln("Content before the bookmark.");
        templateBuilder.StartBookmark("InsertHere");               // Bookmark start.
        templateBuilder.Writeln("[Bookmark placeholder]");        // Placeholder text.
        templateBuilder.EndBookmark("InsertHere");                 // Bookmark end.
        templateBuilder.Writeln("Content after the bookmark.");

        // Save the template so it can be loaded later (optional but follows the rule of creating local files).
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a source document whose content will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        sourceBuilder.Writeln("=== Inserted Document Start ===");
        sourceBuilder.Writeln("This text comes from the source DOCX.");
        sourceBuilder.Writeln("=== Inserted Document End ===");

        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the template, move to the bookmark, and insert the source document.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedTemplate);

        // Position the cursor at the bookmark.
        insertBuilder.MoveToBookmark("InsertHere");

        // Load the source document.
        Document docToInsert = new Document(sourcePath);

        // Insert the source document at the bookmark location.
        insertBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the merged result as PDF.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Simple validation to ensure the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }

        // The program finishes here without waiting for user input.
    }
}
