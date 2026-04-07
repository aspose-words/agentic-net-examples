using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ImageReportExample
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Path to the image that will be inserted into the report.
        public string ImagePath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample image file (a 1x1 pixel PNG) that will be used in the report.
        string imageFile = Path.Combine(outputDir, "sample.png");
        CreateSamplePng(imageFile);

        // 2. Build the DOCX template programmatically.
        string templateFile = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templateFile);

        // 3. Prepare the data model with the image path.
        ReportModel model = new ReportModel { ImagePath = imageFile };

        // 4. Load the template and generate the report using the LINQ Reporting engine.
        Document doc = new Document(templateFile);
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        string reportFile = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportFile);
    }

    // Creates a tiny PNG file from a Base64 string.
    private static void CreateSamplePng(string filePath)
    {
        // Base64 representation of a 1x1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }

    // Generates a DOCX template that contains an image tag inside a textbox.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox that will serve as the container for the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the LINQ Reporting image tag. The tag references the model's ImagePath property.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        doc.Save(filePath);
    }
}
