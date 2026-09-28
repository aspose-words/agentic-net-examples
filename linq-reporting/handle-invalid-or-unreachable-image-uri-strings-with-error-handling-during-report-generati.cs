using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Product
{
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a tiny transparent PNG (1x1 pixel) to use as a valid image.
        string validImagePath = Path.Combine(outputDir, "valid.png");
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
        File.WriteAllBytes(validImagePath, pngBytes);

        // Build data model with one valid and one invalid image path.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new() { Name = "Valid Image", ImagePath = validImagePath },
                new() { Name = "Invalid Image", ImagePath = Path.Combine(outputDir, "nonexistent.jpg") }
            }
        };

        // Replace any missing image with the placeholder so the engine does not throw.
        foreach (var product in model.Products)
        {
            if (!File.Exists(product.ImagePath))
                product.ImagePath = validImagePath;
        }

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");

        // Start a table for each product (header + data row).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 150);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // Return cursor to the table cell after the shape.
        builder.MoveTo(table.LastRow.LastCell.LastParagraph);
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(outputPath);

        // Output status.
        Console.WriteLine(success
            ? "Report generated successfully."
            : "Report generated with errors (see inline messages).");
    }
}
