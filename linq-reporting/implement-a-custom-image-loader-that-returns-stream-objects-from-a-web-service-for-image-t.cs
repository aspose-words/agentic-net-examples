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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a folder for sample assets.
        var assetsFolder = "Assets";
        Directory.CreateDirectory(assetsFolder);

        // Create two tiny PNG images from Base64 strings.
        var pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        var imageBytes = Convert.FromBase64String(pngBase64);
        var imagePath1 = Path.Combine(assetsFolder, "image1.png");
        var imagePath2 = Path.Combine(assetsFolder, "image2.png");
        File.WriteAllBytes(imagePath1, imageBytes);
        File.WriteAllBytes(imagePath2, imageBytes);

        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product
                {
                    Name = "Sample Image 1",
                    ImagePath = imagePath1
                },
                new Product
                {
                    Name = "Sample Image 2",
                    ImagePath = imagePath2
                }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Product Catalog");
        builder.Writeln();

        // Begin foreach over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Start a table for each product.
        var table = builder.StartTable();

        // Name cell.
        builder.InsertCell();
        builder.Write("<<[p.Name]>>");

        // Image cell with a textbox container.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // End row and table.
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        var outputPath = "ProductCatalogReport.docx";
        reportDoc.Save(outputPath);
    }
}

// Root model class.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Product class with a local image path.
public class Product
{
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
