using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Path to the image file that will be inserted into the report.
    public string ImagePath { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image (1x1 red pixel) and save it locally.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will contain the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the image tag with -fitHeight switch to preserve width proportionally.
        builder.Write("<<image [model.ImagePath] -fitHeight>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // Load the template back (required before building the report).
        var loadedTemplate = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            ImagePath = imagePath
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        loadedTemplate.Save(resultPath);
    }
}
