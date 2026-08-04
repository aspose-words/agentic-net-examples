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
        // Create a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample images (1x1 red and blue pixels).
        string imagesDir = Path.Combine(outputDir, "Images");
        Directory.CreateDirectory(imagesDir);
        string redPngPath = Path.Combine(imagesDir, "Red.png");
        string bluePngPath = Path.Combine(imagesDir, "Blue.png");
        CreatePngFromBase64(redPngPath, RedPixelBase64);
        CreatePngFromBase64(bluePngPath, BluePixelBase64);

        // Prepare data for batch reports.
        var reportItems = new List<ReportItem>
        {
            new ReportItem { Title = "First Report", ImagePath = redPngPath },
            new ReportItem { Title = "Second Report", ImagePath = bluePngPath }
        };

        // Build a reusable template document.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // Process each item as an individual report.
        for (int i = 0; i < reportItems.Count; i++)
        {
            // Load the template for each report.
            Document doc = new Document(templatePath);

            // Use the single ReportItem as the root data source.
            ReportItem model = reportItems[i];

            // Build the report using LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, $"Report_{i + 1}.docx");
            doc.Save(reportPath);
        }
    }

    // Creates a simple template with a title and an image inside a textbox.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Batch Report Example");

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 150);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("Title: <<[model.Title]>>");
        builder.Writeln("<<image [model.ImagePath] -fitSize>>");

        doc.Save(filePath);
    }

    // Writes a PNG file from a Base64 string.
    private static void CreatePngFromBase64(string filePath, string base64)
    {
        byte[] bytes = Convert.FromBase64String(base64);
        File.WriteAllBytes(filePath, bytes);
    }

    // Model class for a single report item.
    public class ReportItem
    {
        public string Title { get; set; } = "";
        public string ImagePath { get; set; } = "";
    }

    // Base64 for a 1x1 red pixel PNG.
    private const string RedPixelBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6WQAAAAASUVORK5CYII=";

    // Base64 for a 1x1 blue pixel PNG.
    private const string BluePixelBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AAAwMCAO+X6WQAAAAASUVORK5CYII=";
}
