using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class ReportModel
{
    // URI or file path to the image. Initialized to avoid nullable warnings.
    public string ImageUri { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple 1x1 pixel PNG image from a Base64 string.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YVbZc8AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string validImagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(validImagePath, pngBytes);

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Image Report");

        // Insert a textbox that will host the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // LINQ Reporting tag that inserts an image from the model.
        builder.Write("<<image [model.ImageUri] -fitSize>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Generate a report with a valid image URI.
        // -----------------------------------------------------------------
        var reportValid = new Document(templatePath);
        var modelValid = new ReportModel { ImageUri = validImagePath };

        var engine = new ReportingEngine
        {
            // Inline error messages will be inserted if the image cannot be loaded.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        bool successValid = engine.BuildReport(reportValid, modelValid, "model");
        string validReportPath = Path.Combine(outputDir, "report_valid.docx");
        reportValid.Save(validReportPath);

        // -----------------------------------------------------------------
        // Step 3: Generate a report with an invalid/unreachable image URI.
        // -----------------------------------------------------------------
        var reportInvalid = new Document(templatePath);
        var modelInvalid = new ReportModel { ImageUri = Path.Combine(outputDir, "nonexistent.png") };

        bool successInvalid;
        try
        {
            // BuildReport will return false when an error occurs and InlineErrorMessages is set.
            successInvalid = engine.BuildReport(reportInvalid, modelInvalid, "model");
        }
        catch (Exception ex)
        {
            // If an unexpected exception occurs, treat the build as failed.
            Console.WriteLine($"Unexpected error while building report: {ex.Message}");
            successInvalid = false;
        }

        string invalidReportPath = Path.Combine(outputDir, "report_invalid.docx");
        reportInvalid.Save(invalidReportPath);

        // Output the results.
        Console.WriteLine($"Valid image report generated: {successValid} -> {validReportPath}");
        Console.WriteLine($"Invalid image report generated: {successInvalid} -> {invalidReportPath}");
    }
}
