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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string workDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(workDir, "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image (1x1 pixel) from a Base64 string.
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XG6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the data model and load the image bytes.
        ReportModel model = new ReportModel
        {
            ImageUri = imagePath,
            ImageData = pngBytes
        };

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will hold the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 250, 250);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Public data model class.
public class ReportModel
{
    // Original image URI (kept for reference).
    public string ImageUri { get; set; } = string.Empty;

    // Image data as a byte array that will be inserted into the report.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
