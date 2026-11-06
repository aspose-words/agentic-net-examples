using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string templatePath = "Template.docx";
        string sourcePath = "Source.docx";
        string outputPath = "Result.html";

        // -------------------------------------------------
        // Create a styled template document (DOCX)
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(templateDoc);

        // Add a title using the built‑in Title style
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        templateBuilder.Writeln("Template Title");

        // Add normal paragraph content
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        templateBuilder.Writeln("This is the template content before insertion.");

        // Save the template to disk
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // -------------------------------------------------
        // Create a source document (DOCX) to be inserted
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // Add some content with a different style
        sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        sourceBuilder.Writeln("Source Document Heading");

        sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        sourceBuilder.Writeln("This is the source document content that will be inserted.");

        // Save the source document to disk
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // Load the template and insert the source document
        // using ImportFormatMode.UseDestinationStyles
        // -------------------------------------------------
        Document destination = new Document(templatePath);
        DocumentBuilder destBuilder = new DocumentBuilder(destination);

        // Load the source document (demonstrates the workflow)
        Document toInsert = new Document(sourcePath);

        // Insert the source document at the current builder position
        destBuilder.InsertDocument(toInsert, ImportFormatMode.UseDestinationStyles);

        // -------------------------------------------------
        // Save the combined document as HTML
        // -------------------------------------------------
        destination.Save(outputPath, SaveFormat.Html);

        // -------------------------------------------------
        // Validate that the output file was created
        // -------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }

        // Optional quick validation: ensure the HTML contains text from both documents
        string htmlContent = File.ReadAllText(outputPath);
        if (!htmlContent.Contains("Template Title") ||
            !htmlContent.Contains("Source Document Heading"))
        {
            throw new InvalidOperationException("The merged HTML does not contain expected content.");
        }

        // Clean up temporary files (optional)
        File.Delete(templatePath);
        File.Delete(sourcePath);
    }
}
