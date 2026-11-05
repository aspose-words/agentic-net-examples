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
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string reportPath = Path.Combine(workDir, "report.docx");
        string imagePath = Path.Combine(workDir, "sample.png");

        // Create a simple PNG image file (1x1 red pixel) without using System.Drawing
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXK5wAAAABJRU5ErkJggg==";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Image Report");
        builder.Writeln();

        // Insert a textbox that will hold the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 250, 150);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template (optional, can use the same instance)
        Document doc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            Title = "Sample Image",
            ImageUri = imagePath,
            ImageData = File.ReadAllBytes(imagePath)
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(reportPath);
    }
}

// Data model used by the LINQ reporting engine
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string ImageUri { get; set; } = string.Empty;
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
