using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(baseDir, "Template.docx");
        string sourcePath = Path.Combine(baseDir, "Source.docx");
        string resultPdfPath = Path.Combine(baseDir, "Result.pdf");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a bookmark named "InsertHere".
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        templateBuilder.Writeln("This is the template document.");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("[Placeholder for inserted content]");
        templateBuilder.EndBookmark("InsertHere");
        templateBuilder.Writeln("End of template.");

        // Save the template as DOCX.
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a source document that will be inserted at the bookmark.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        sourceBuilder.Writeln("=== Inserted Document Start ===");
        sourceBuilder.Writeln("This content comes from the source document.");
        sourceBuilder.Writeln("=== Inserted Document End ===");

        // Save the source as DOCX.
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the template and source documents.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Document loadedSource = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 4. Move to the bookmark and insert the source document.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedTemplate);
        insertBuilder.MoveToBookmark("InsertHere");

        // Insert the source document at the bookmark position, preserving its formatting.
        insertBuilder.InsertDocument(loadedSource, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Save the merged result as PDF.
        // -----------------------------------------------------------------
        loadedTemplate.Save(resultPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 6. Simple validation to ensure the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPdfPath))
        {
            throw new FileNotFoundException("The resulting PDF was not created.", resultPdfPath);
        }

        // Optional: Verify that the merged document contains text from the source.
        string mergedText = loadedTemplate.GetText();
        if (!mergedText.Contains("Inserted Document Start"))
        {
            throw new InvalidOperationException("The source content was not inserted correctly.");
        }

        // Program ends without waiting for user input.
    }
}
