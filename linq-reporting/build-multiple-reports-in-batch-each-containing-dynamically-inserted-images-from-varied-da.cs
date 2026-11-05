using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ReportItem
{
    public string Title { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare directories.
        string imagesDir = "images";
        string outputDir = "output";
        Directory.CreateDirectory(imagesDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PNG images from Base64 strings.
        var redPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAFklEQVR42mNk+M9Qz0AEYBxVSFIAABcAAf+XKc8AAAAASUVORK5CYII="; // 16x16 red square
        var bluePngBase64 = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAFklEQVR42mNkYGBg+M+ABYwMDAwAAJcAAf+K6V8AAAAASUVORK5CYII="; // 16x16 blue square

        string redImagePath = Path.Combine(imagesDir, "red.png");
        string blueImagePath = Path.Combine(imagesDir, "blue.png");
        File.WriteAllBytes(redImagePath, Convert.FromBase64String(redPngBase64));
        File.WriteAllBytes(blueImagePath, Convert.FromBase64String(bluePngBase64));

        // Sample data for batch reports.
        var reports = new List<ReportItem>
        {
            new() { Title = "Red Image Report", ImagePath = redImagePath },
            new() { Title = "Blue Image Report", ImagePath = blueImagePath }
        };

        // Create a reusable template document.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title placeholder.
        builder.Writeln("Report for: <<[item.Title]>>");

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [item.ImagePath] -fitSize>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Build each report using the template.
        var engine = new ReportingEngine();

        foreach (var report in reports)
        {
            var doc = new Document(templatePath);
            engine.BuildReport(doc, report, "item");

            string outputPath = Path.Combine(outputDir, $"{report.Title.Replace(' ', '_')}.docx");
            doc.Save(outputPath);
        }
    }
}
