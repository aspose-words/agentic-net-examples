using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define file paths
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.html");

        // Create the template document with a heading style
        Document templateDoc = new();
        DocumentBuilder templateBuilder = new(templateDoc);
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        templateBuilder.Writeln("Template Title");
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        templateBuilder.Writeln("This is the template content.");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // Create the source document with a different heading style
        Document sourceDoc = new();
        DocumentBuilder sourceBuilder = new(sourceDoc);
        sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        sourceBuilder.Writeln("Source Title");
        sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        sourceBuilder.Writeln("This is the source content with a different style.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the documents from disk (simulating real-world scenario)
        Document loadedTemplate = new(templatePath);
        Document loadedSource = new(sourcePath);

        // Insert the source document into the template using destination styles
        DocumentBuilder insertBuilder = new(loadedTemplate);
        // Move cursor to the end of the template document
        insertBuilder.MoveToDocumentEnd();
        insertBuilder.InsertDocument(loadedSource, ImportFormatMode.UseDestinationStyles);

        // Save the combined document as HTML
        loadedTemplate.Save(outputPath, SaveFormat.Html);

        // Validation: ensure output file exists
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Output file was not created: {outputPath}");
        }

        // Validation: ensure both source markers are present in the HTML
        string htmlContent = File.ReadAllText(outputPath);
        if (!htmlContent.Contains("Template Title") || !htmlContent.Contains("Source Title"))
        {
            throw new InvalidOperationException("Combined HTML does not contain expected content from both documents.");
        }

        // Indicate successful execution (no interactive prompts)
        Console.WriteLine("Document merged and saved as HTML successfully.");
    }
}
