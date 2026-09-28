using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare directories
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");

        // Create sample images
        string imagePath1 = Path.Combine(Directory.GetCurrentDirectory(), "image1.png");
        string imagePath2 = Path.Combine(Directory.GetCurrentDirectory(), "image2.png");
        WritePngFromBase64(imagePath1, "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/6ZkZAAAAAElFTkSuQmCC"); // red
        WritePngFromBase64(imagePath2, "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8DAAQAD/6X+6wAAAABJRU5ErkJggg=="); // green

        // Build template document with an image placeholder inside a textbox
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title: <<[Item.Title]>>");
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [Item.ImagePath] -fitSize>>");
        templateDoc.Save(templatePath);

        // Sample data for batch reports
        var items = new List<ReportItem>
        {
            new ReportItem { Title = "FirstReport", ImagePath = imagePath1 },
            new ReportItem { Title = "SecondReport", ImagePath = imagePath2 }
        };

        // Optional: enable reflection optimization
        ReportingEngine.UseReflectionOptimization = true;

        // Generate each report
        foreach (var item in items)
        {
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, item, "Item");
            string outPath = Path.Combine(outputDir, $"{item.Title}.docx");
            doc.Save(outPath);
        }
    }

    // Helper to write a PNG file from a Base64 string
    private static void WritePngFromBase64(string filePath, string base64)
    {
        byte[] bytes = Convert.FromBase64String(base64);
        File.WriteAllBytes(filePath, bytes);
    }
}

// Public data model for a single report
public class ReportItem
{
    public string Title { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
