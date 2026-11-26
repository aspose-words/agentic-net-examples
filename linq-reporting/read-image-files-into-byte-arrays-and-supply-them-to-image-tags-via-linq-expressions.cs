using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    // Public data model for the report.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public byte[] ImageData { get; set; } = Array.Empty<byte>();
    }

    // Wrapper class required by the LINQ Reporting engine (no anonymous types).
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a tiny PNG image (1x1 red pixel) from a Base64 string.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Sample data.
        var products = new List<Product>
        {
            new Product { Name = "Red Pixel", ImageData = File.ReadAllBytes(imagePath) },
            new Product { Name = "Another Red Pixel", ImageData = File.ReadAllBytes(imagePath) }
        };
        var model = new ReportModel { Products = products };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table for each product.
        Table table = builder.StartTable();

        // First cell: product name.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        // Second cell: image inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageData] -fitSize>>");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
