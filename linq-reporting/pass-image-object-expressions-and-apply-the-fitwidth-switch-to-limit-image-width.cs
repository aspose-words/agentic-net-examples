using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Image data as a byte array – supported by LINQ Reporting image tag.
    public byte[] ImageData { get; }

    public ReportModel(byte[] imageData)
    {
        ImageData = imageData ?? throw new ArgumentNullException(nameof(imageData));
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple 1x1 PNG image (transparent) from a Base64 string.
        // This avoids using System.Drawing types which may be unavailable.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXcAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Save the image file so we can see the result on disk (optional).
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Build the data model.
        var model = new ReportModel(imageBytes);

        // Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitWidth>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Generate the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(outputDir, "output.docx");
        reportDoc.Save(outputPath);
    }
}
