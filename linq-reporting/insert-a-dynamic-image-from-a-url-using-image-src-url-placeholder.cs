using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // ---------- Create template ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag that will receive a byte[] with the image data.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // ---------- Prepare data ----------
        var model = new ReportModel
        {
            // Use a small embedded PNG image (1x1 transparent pixel) to avoid network calls.
            ImageData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=")
        };

        // ---------- Build report ----------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Public data model used by the template.
public class ReportModel
{
    // Image data as a byte array; the image tag can consume this type.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
