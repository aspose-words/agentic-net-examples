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
        string outputPath = "Result.pdf";

        // Create a template document with a bookmark named "InsertHere"
        var templateDoc = new Document();
        var templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.Writeln("This is the beginning of the template document.");
        templateBuilder.StartBookmark("InsertHere");
        templateBuilder.Writeln("[Content will be inserted here]");
        templateBuilder.EndBookmark("InsertHere");
        templateBuilder.Writeln("This is the end of the template document.");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // Create a source document that will be inserted
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("=== Start of Source Document ===");
        sourceBuilder.Writeln("This is some sample content from the source document.");
        sourceBuilder.Writeln("=== End of Source Document ===");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the template and source documents
        var mainDoc = new Document(templatePath);
        var insertDoc = new Document(sourcePath);

        // Move to the bookmark and insert the source document
        var builder = new DocumentBuilder(mainDoc);
        builder.MoveToBookmark("InsertHere");
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF
        mainDoc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("The merged PDF was not created successfully.");
        }

        // Clean up temporary files (optional)
        File.Delete(templatePath);
        File.Delete(sourcePath);
    }
}
