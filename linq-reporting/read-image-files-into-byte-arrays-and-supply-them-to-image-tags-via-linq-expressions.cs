using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create two 1x1 pixel PNG images (red and green) from Base64 strings.
        string[] imagePaths =
        {
            Path.Combine(outputDir, "red.png"),
            Path.Combine(outputDir, "green.png")
        };

        string[] base64Images =
        {
            // Red pixel PNG
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXcVwAAAABJRU5ErkJggg==",
            // Green pixel PNG
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8DAAQAB/6F9WQAAAABJRU5ErkJggg=="
        };

        for (int i = 0; i < imagePaths.Length; i++)
        {
            byte[] pngBytes = Convert.FromBase64String(base64Images[i]);
            File.WriteAllBytes(imagePaths[i], pngBytes);
        }

        // Load images into the data model.
        var products = new List<Product>();
        for (int i = 0; i < imagePaths.Length; i++)
        {
            products.Add(new Product
            {
                Name = $"Product {(i + 1)}",
                ImageBytes = File.ReadAllBytes(imagePaths[i])
            });
        }

        var model = new ReportModel
        {
            Products = products
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Catalog");

        // Begin foreach loop.
        builder.Writeln("<<foreach [p in Products]>>");

        // Start table inside the loop.
        Table table = builder.StartTable();

        // Header row (will be repeated for each product – acceptable for this simple example).
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row for each product.
        builder.InsertCell();
        builder.Write("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageBytes] -fitSize>>");

        builder.EndRow();

        // Close the table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        templateDoc.Save(reportPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public byte[] ImageBytes { get; set; } = Array.Empty<byte>();
}
