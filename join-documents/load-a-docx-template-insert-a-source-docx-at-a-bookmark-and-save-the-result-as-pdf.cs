using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files
        string templatePath = "Template.docx";
        string sourcePath = "Source.docx";
        string outputPdfPath = "Result.pdf";

        // Create sample documents (template and source)
        CreateTemplateDocument(templatePath);
        CreateSourceDocument(sourcePath);

        // Load the template and source documents
        Document templateDoc = new Document(templatePath);
        Document sourceDoc = new Document(sourcePath);

        // Verify that the bookmark exists in the template
        if (templateDoc.Range.Bookmarks["InsertHere"] == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found in the template.");

        // Insert the source document at the bookmark location
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.MoveToBookmark("InsertHere");
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF
        templateDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created successfully
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the merged PDF output.");

        // Optional cleanup (comment out if you want to keep the sample files)
        // File.Delete(templatePath);
        // File.Delete(sourcePath);
    }

    // Creates a simple DOCX template with a bookmark named "InsertHere"
    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is the template document.");
        builder.StartBookmark("InsertHere");
        builder.Writeln("[Content will be inserted here]");
        builder.EndBookmark("InsertHere");
        builder.Writeln("End of template.");

        doc.Save(path, SaveFormat.Docx);
    }

    // Creates a simple DOCX source document that will be inserted into the template
    private static void CreateSourceDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is the source document that will be inserted.");
        builder.Writeln("It contains multiple paragraphs.");
        builder.Writeln("End of source document.");

        doc.Save(path, SaveFormat.Docx);
    }
}
