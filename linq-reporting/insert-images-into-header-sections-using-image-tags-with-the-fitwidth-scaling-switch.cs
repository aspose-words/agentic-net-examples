using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory("Output");

        // Create a simple 1x1 PNG image and save it locally.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK0cAAAAASUVORK5CYII=");
        string imagePath = Path.Combine("Output", "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Define the data model.
        ReportModel model = new ReportModel
        {
            ImagePath = Path.GetFullPath(imagePath)
        };

        // Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a textbox shape that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with fitWidth scaling switch.
        builder.Write("<<image [model.ImagePath] -fitWidth>>");

        // Add some body content (optional).
        builder.MoveToDocumentEnd();
        builder.Writeln("Report generated with header image.");

        // Save the template to disk.
        string templatePath = Path.Combine("Output", "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(report, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine("Output", "Report.docx");
        report.Save(reportPath);
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    // Path to the image file that will be inserted into the header.
    public string ImagePath { get; set; } = string.Empty;
}
