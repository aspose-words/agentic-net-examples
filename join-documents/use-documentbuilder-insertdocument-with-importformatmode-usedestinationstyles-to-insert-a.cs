using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string templatePath = "Template.docx";
        string sourcePath = "Source.docx";
        string resultPath = "Result.html";

        // Create a styled template document
        var templateDoc = new Document();
        var templateBuilder = new DocumentBuilder(templateDoc);
        templateBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        templateBuilder.Writeln("Template Title");
        templateBuilder.Writeln("This is the template body.");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // Create a source DOCX document to be inserted
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        sourceBuilder.Writeln("Source content paragraph 1.");
        sourceBuilder.Writeln("Source content paragraph 2.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the template and source documents
        var destination = new Document(templatePath);
        var source = new Document(sourcePath);

        // Insert the source document into the template using destination styles
        var destBuilder = new DocumentBuilder(destination);
        destBuilder.MoveToDocumentEnd();
        destBuilder.InsertDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the merged document as HTML
        destination.Save(resultPath, SaveFormat.Html);

        // Validate that the HTML file was created
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {resultPath}");
        }

        // Simple validation that both source and template contents are present in the HTML
        string htmlContent = File.ReadAllText(resultPath);
        if (!htmlContent.Contains("Template Title") || !htmlContent.Contains("Source content paragraph 1"))
        {
            throw new InvalidOperationException("The merged HTML does not contain expected content from both documents.");
        }

        // Clean up temporary files (optional)
        // File.Delete(templatePath);
        // File.Delete(sourcePath);
    }
}
