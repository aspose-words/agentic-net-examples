using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Base64Image { get; set; } = string.Empty;

    public byte[] ImageBytes => Convert.FromBase64String(Base64Image);
}

public class Program
{
    public static void Main()
    {
        // Prepare sample Base64 image (1x1 red PNG).
        const string sampleBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";

        // Create data model.
        var model = new ReportModel { Base64Image = sampleBase64 };

        // Create template document.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a textbox to hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageBytes] -fitSize>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model.
        engine.BuildReport(reportDoc, model, "model");

        // Ensure output directory exists.
        Directory.CreateDirectory("output");

        // Save the generated report.
        var outputPath = Path.Combine("output", "ReportWithImage.docx");
        reportDoc.Save(outputPath);
    }
}
