using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a simple PNG image (1x1 pixel, red) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAocB9pVYB2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Ensure a folder for the template and output exists.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string reportPath = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple title using a model property.
        builder.Writeln("<<[model.Title]>>");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag expects a byte[] expression; -fitSize scales the image to the textbox.
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Save the template to disk (required before BuildReport).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model supplying the image bytes.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Title = "Sample Image Report",
            ImageData = pngBytes
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Title displayed above the image.
    public string Title { get; set; } = string.Empty;

    // Image data supplied as a byte array.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
