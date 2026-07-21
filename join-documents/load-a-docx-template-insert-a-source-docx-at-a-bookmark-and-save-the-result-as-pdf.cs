using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(baseDir, "Template.docx");
        string sourcePath = Path.Combine(baseDir, "Source.docx");
        string outputPdfPath = Path.Combine(baseDir, "MergedResult.pdf");

        // -----------------------------------------------------------------
        // Create a template document with a bookmark where the source will be inserted.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("=== Template Document Start ===");
        tmplBuilder.StartBookmark("InsertHere");
        tmplBuilder.Writeln("[Placeholder for source document]");
        tmplBuilder.EndBookmark("InsertHere");
        tmplBuilder.Writeln("=== Template Document End ===");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a source document that will be inserted into the template.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln(">>> This is the source document content. <<<");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents from disk.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        Document loadedSource = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Insert the source document at the bookmark.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedTemplate);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(loadedSource, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged document as PDF.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF file exists and contains text from both documents.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The merged PDF file was not created.");

        // Load the PDF back as a Word document to verify its text content.
        Document verificationDoc = new Document(outputPdfPath);
        string mergedText = verificationDoc.GetText();

        if (!mergedText.Contains("=== Template Document Start ===") ||
            !mergedText.Contains(">>> This is the source document content. <<<") ||
            !mergedText.Contains("=== Template Document End ==="))
        {
            throw new InvalidOperationException("The merged PDF does not contain expected content.");
        }

        // Optional: indicate success (no interactive I/O required).
        Console.WriteLine("Document merged and saved as PDF successfully.");
    }
}
