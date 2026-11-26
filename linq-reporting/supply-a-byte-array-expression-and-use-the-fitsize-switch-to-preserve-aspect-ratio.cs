using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Byte array containing PNG image data.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();

    public ReportModel()
    {
        // A 1x1 pixel red PNG image encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        ImageData = Convert.FromBase64String(base64Png);
    }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox to host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // Image tag using a byte array expression with -fitSize switch to preserve aspect ratio.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
