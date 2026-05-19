using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Holds image data as a byte array that the LINQ reporting engine can use.
    public byte[] Image { get; set; } = Array.Empty<byte>();

    // Additional sample data.
    public string Title { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple 1x1 PNG image in memory (no System.Drawing usage).
        // -----------------------------------------------------------------
        // This is a minimal valid PNG (transparent pixel) encoded in Base64.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // -----------------------------------------------------------------
        // 2. Build the data model for the report.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Image = imageBytes,
            Title = "Sample Image Report"
        };

        // -----------------------------------------------------------------
        // 3. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional title paragraph.
        builder.Writeln($"<<[model.Title]>>");
        builder.Writeln();

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // The image tag expects an expression that evaluates to a supported image type (byte[] here).
        builder.Write("<<image [model.Image] -fitSize>>");

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "ReportWithImage.docx");
        doc.Save(resultPath);
    }
}
