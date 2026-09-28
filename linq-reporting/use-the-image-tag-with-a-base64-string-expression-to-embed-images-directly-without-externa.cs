using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Byte array containing the image data.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}

public class Program
{
    public static void Main()
    {
        // Sample 1x1 PNG image (transparent) as a byte array.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6ZcAAAAASUVORK5CYII=");

        // Prepare the model with the image data.
        ReportModel model = new() { ImageData = pngBytes };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag using the byte array expression.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
