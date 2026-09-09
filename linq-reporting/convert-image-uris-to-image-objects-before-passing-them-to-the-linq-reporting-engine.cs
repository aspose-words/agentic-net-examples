using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables; // Needed for Table type

#nullable enable

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        const string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image that will be referenced by URI.
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        CreateSampleImage(sampleImagePath);

        // Build the data model.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Product A", ImageUri = sampleImagePath },
                new Product { Name = "Product B", ImageUri = sampleImagePath }
            }
        };

        // Convert image URIs to byte[] objects before reporting.
        foreach (var product in model.Products)
        {
            // Load the image bytes from the file system.
            product.ImageData = File.ReadAllBytes(product.ImageUri);
        }

        // Create the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach block over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: product name and image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row (repeated for each product).
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox to host the image tag (required by LINQ Reporting).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageData] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outputPath);
    }

    // Helper method to create a simple 1x1 PNG image.
    private static void CreateSampleImage(string path)
    {
        // This is a minimal 1x1 pixel transparent PNG.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/AL+XKXK" +
            "AAAAAElFTkSuQmCC";

        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }
}

// Root data model passed to the reporting engine.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Individual product with name and image data.
public class Product
{
    // Display name.
    public string Name { get; set; } = string.Empty;

    // Original image URI (file path).
    public string ImageUri { get; set; } = string.Empty;

    // Image data used by the <<image>> tag (byte array).
    public byte[]? ImageData { get; set; }
}
