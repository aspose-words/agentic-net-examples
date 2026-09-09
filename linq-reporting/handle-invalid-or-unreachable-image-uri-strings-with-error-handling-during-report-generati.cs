using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a tiny red PNG image (1x1 pixel) and save it locally.
        string validImagePath = Path.Combine(outputDir, "valid.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
        File.WriteAllBytes(validImagePath, pngBytes);

        // Prepare the data model.
        var model = new ReportModel
        {
            Title = "Image URI Error Handling Demo",
            Items = new List<ReportItem>
            {
                new ReportItem { ImageUri = validImagePath },                                 // Valid local file.
                new ReportItem { ImageUri = "https://example.com/missing.jpg" }, // Invalid remote URI.
                new ReportItem { ImageUri = @"C:\nonexistent\image.png" }      // Invalid local path.
            }
        };

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in model.Items]>>");

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag inside the textbox. Use -fitSize to keep original dimensions.
        builder.Write("<<image [item.ImageUri] -fitSize>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "ReportWithImages.docx");
        doc.Save(outputPath);

        // Output the result.
        Console.WriteLine($"Report generation success flag: {success}");
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// Root data model.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<ReportItem> Items { get; set; } = new();
}

// Item containing an image URI (could be a file path or a web URL).
public class ReportItem
{
    public string ImageUri { get; set; } = string.Empty;
}
