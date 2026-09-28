using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportModel
{
    // Byte array containing image data (a simple 1x1 PNG).
    public byte[] ImageData { get; set; } = Array.Empty<byte>();

    public ReportModel()
    {
        // Base64‑encoded 1x1 pixel PNG (transparent).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        ImageData = Convert.FromBase64String(base64Png);
    }
}

public class Program
{
    public static void Main()
    {
        // ---------- Create template document ----------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // Image tag using byte array expression with -fitSize switch to preserve aspect ratio.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // ---------- Load template and build report ----------
        var doc = new Document(templatePath);

        // Prepare data model.
        var model = new ReportModel();

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
