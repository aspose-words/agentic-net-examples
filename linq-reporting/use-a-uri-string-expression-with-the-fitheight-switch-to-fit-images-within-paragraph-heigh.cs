using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a simple 1x1 pixel PNG image.
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox to host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting image tag with fitHeight switch.
        builder.Write("<<image [model.ImageUri] -fitHeight>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            ImageUri = Path.GetFullPath(imagePath) // URI string expression.
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);
    }
}

// Public data model for the report.
public class ReportModel
{
    // Path or URI to the image file.
    public string ImageUri { get; set; } = string.Empty;
}
