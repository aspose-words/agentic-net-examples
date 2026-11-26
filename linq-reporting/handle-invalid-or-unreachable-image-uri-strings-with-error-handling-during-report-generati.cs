using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample image file (a 1x1 pixel PNG)
        string validImagePath = "sample.png";
        CreateSamplePng(validImagePath);

        // Prepare data model with one valid and one invalid image path
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Valid Image", ImagePath = validImagePath },
                new Item { Name = "Invalid Image", ImagePath = "nonexistent.png" }
            }
        };

        // Create template document programmatically
        Document template = new();
        DocumentBuilder builder = new(template);

        // Write a simple heading
        builder.Writeln("Image Report");
        builder.Writeln();

        // Begin foreach loop over Items
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>");

        // Insert a textbox to hold the image
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 150);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [item.ImagePath] -fitSize>>");

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report)
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for report generation
        Document report = new(templatePath);

        // Configure reporting engine to embed inline error messages
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(report, model, "model");

        // Save the generated report
        string outputPath = "output.docx";
        report.Save(outputPath);

        // Output result status
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }

    // Creates a minimal 1x1 PNG file at the specified path
    private static void CreateSamplePng(string path)
    {
        // PNG file bytes for a 1x1 transparent pixel
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        File.WriteAllBytes(path, pngBytes);
    }
}

// Data model classes
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
