using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create template document with an image placeholder inside a textbox
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox to host the image
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Prepare model with Base64 image string and converted byte array
        ReportModel model = new()
        {
            Base64Image = SampleBase64Png,
            ImageData = Convert.FromBase64String(SampleBase64Png)
        };

        // Build the report
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);
    }

    // Sample 1x1 pixel PNG encoded in Base64
    private const string SampleBase64Png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
        "Zc6cAAAAASUVORK5CYII=";

    public class ReportModel
    {
        // Base64 string (not used directly in the template but kept for completeness)
        public string Base64Image { get; set; } = string.Empty;

        // Byte array used by the image tag
        public byte[] ImageData { get; set; } = Array.Empty<byte>();
    }
}
