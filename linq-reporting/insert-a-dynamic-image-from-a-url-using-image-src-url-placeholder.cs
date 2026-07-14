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

        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image file locally
        string imagePath = Path.Combine(outputDir, "sample.png");
        WriteSamplePng(imagePath);

        // Create the template document with a textbox containing the image tag
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Prepare the model
        var model = new ReportModel { ImagePath = imagePath };

        // Load the template and build the report
        Document template = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "report.docx");
        template.Save(reportPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        doc.Save(path);
    }

    private static void WriteSamplePng(string filePath)
    {
        // A 1x1 pixel transparent PNG (base64 encoded)
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XbZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }

    public class ReportModel
    {
        // Path to the image file to be inserted
        public string ImagePath { get; set; } = string.Empty;
    }
}
