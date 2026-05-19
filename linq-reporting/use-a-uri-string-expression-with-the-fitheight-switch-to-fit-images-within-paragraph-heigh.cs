using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string workDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(workDir, "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple 1x1 red PNG image and save it to disk.
        string imagePath = Path.Combine(outputDir, "red.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+VbV8AAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // Build the LINQ Reporting template.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph describing the image.
        builder.Writeln("Image fitted to paragraph height:");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
        builder.MoveTo(textBox.FirstParagraph);

        // Insert the image tag. The expression refers to the model property.
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Data model exposing the image URI.
        ReportModel model = new ReportModel(imagePath);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report using the model as the root data source named "model".
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // Save the final document.
        string resultPath = Path.Combine(outputDir, "result.docx");
        loadedTemplate.Save(resultPath);
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path or URI to the image that will be inserted.
    public string ImageUri { get; set; } = string.Empty;

    public ReportModel(string imageUri)
    {
        ImageUri = imageUri;
    }
}
