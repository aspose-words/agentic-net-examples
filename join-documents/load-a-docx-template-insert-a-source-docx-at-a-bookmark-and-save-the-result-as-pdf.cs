using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the template, source document, and final PDF.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultPdfPath = Path.Combine(outputDir, "Result.pdf");

        // ---------- Create a template document with a bookmark ----------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is the template document.");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("Bookmark placeholder.");
        templateBuilder.EndBookmark("InsertHere");
        templateBuilder.Writeln("End of template.");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // ---------- Create a source document that will be inserted ----------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the inserted source document.");
        sourceBuilder.Writeln("Additional source content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Load the template and insert the source at the bookmark ----------
        Document mainDoc = new Document(templatePath);
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.MoveToBookmark("InsertHere");

        Document docToInsert = new Document(sourcePath);
        // Insert the source document at the current cursor position, preserving its formatting.
        mainBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged document as PDF ----------
        mainDoc.Save(resultPdfPath, SaveFormat.Pdf);

        // ---------- Validate that the PDF was created ----------
        if (!File.Exists(resultPdfPath))
        {
            throw new InvalidOperationException("The PDF file was not created.");
        }

        // Optional: indicate success.
        Console.WriteLine($"Merged PDF created at: {resultPdfPath}");
    }
}
