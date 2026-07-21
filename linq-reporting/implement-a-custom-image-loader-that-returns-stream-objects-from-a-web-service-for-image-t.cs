using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a tiny PNG image (1x1 pixel, red) and save it locally.
        const string imageFileName = "sample.png";
        CreateSamplePng(imageFileName);

        // Build the template document.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox shape to host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag referencing a string property on the model.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template.
        doc.Save(templatePath);

        // Prepare the model with the local image path.
        var model = new ReportModel
        {
            ImagePath = Path.GetFullPath(imageFileName)
        };

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    // Creates a simple 1x1 red PNG file from an embedded base64 string.
    private static void CreateSamplePng(string filePath)
    {
        // Base64 for a 1x1 red PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }
}

// Model class exposing an ImagePath property that points to a local image file.
public class ReportModel
{
    // Full path to the image file to be inserted.
    public string ImagePath { get; set; } = string.Empty;
}
