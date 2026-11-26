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
        // Prepare output folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        string imagesDir = Path.Combine(outputDir, "images");
        Directory.CreateDirectory(imagesDir);

        // Create sample 1x1 PNG images
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YVhVfUAAAAASUVORK5CYII=");
        string imagePath1 = Path.Combine(imagesDir, "product1.png");
        string imagePath2 = Path.Combine(imagesDir, "product2.png");
        File.WriteAllBytes(imagePath1, pngData);
        File.WriteAllBytes(imagePath2, pngData);

        // -----------------------------------------------------------------
        // Build the template document programmatically
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Products
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: Name and Image
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Product Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row (repeated for each product)
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox that will hold the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // End the data row
        builder.EndRow();

        // Finish the table and the foreach block
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new()
        {
            Products = new List<Product>
            {
                new Product { Name = "Product 1", ImagePath = imagePath1 },
                new Product { Name = "Product 2", ImagePath = imagePath2 }
            }
        };

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
}
