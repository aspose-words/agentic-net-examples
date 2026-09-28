using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code pages (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare a sample image file
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        if (!File.Exists(imagePath))
        {
            // 1x1 transparent PNG
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Create the template document
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a chart (valid placement)
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Move to the end of the document (outside the chart) and insert valid LINQ Reporting tags
        builder.MoveToDocumentEnd();
        builder.Writeln("<<image [model.ImagePath]>>");
        builder.Writeln("<<bookmark [model.BookmarkName]>>Bookmarked Content<</bookmark>>");
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Save the template (optional, demonstrates persistence)
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare model data
        ReportModel model = new()
        {
            ImagePath = imagePath,
            BookmarkName = "SampleBookmark",
            Url = "https://example.com",
            LinkText = "Example Site"
        };

        // Build the report with inline error messages to capture validation errors
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the resulting document
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        // Output validation result
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}

// Data model used by the LINQ Reporting engine
public class ReportModel
{
    public string ImagePath { get; set; } = "";
    public string BookmarkName { get; set; } = "";
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}
