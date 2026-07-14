using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for the Table class

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
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string imagesDir = Path.Combine(workDir, "Images");
        Directory.CreateDirectory(imagesDir);
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create two tiny PNG files to act as product images.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XG8cAAAAASUVORK5CYII=");
        string imgPath1 = Path.Combine(imagesDir, "product1.png");
        string imgPath2 = Path.Combine(imagesDir, "product2.png");
        File.WriteAllBytes(imgPath1, pngBytes);
        File.WriteAllBytes(imgPath2, pngBytes);

        // Build the data model.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Product A", ImagePath = imgPath1 },
                new Product { Name = "Product B", ImagePath = imgPath2 }
            }
        };

        // Create the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Catalog");
        builder.Writeln();

        // Begin the LINQ Reporting foreach block.
        builder.Writeln("<<foreach [p in Products]>>");

        // Table with two columns: Name and Image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row – will be repeated for each product.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox and place the image tag inside it.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 150);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        doc.Save(templatePath);

        // Build the final report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}
