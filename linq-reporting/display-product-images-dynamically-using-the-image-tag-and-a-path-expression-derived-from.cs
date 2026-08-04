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
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample image files
        string img1Path = Path.Combine(outputDir, "product1.png");
        string img2Path = Path.Combine(outputDir, "product2.png");

        // 1x1 red PNG
        byte[] redPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6V6eUAAAAASUVORK5CYII=");
        // 1x1 blue PNG
        byte[] bluePng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/6cKcVwAAAABJRU5ErkJggg==");

        File.WriteAllBytes(img1Path, redPng);
        File.WriteAllBytes(img2Path, bluePng);

        // Prepare data model
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Red Product", ImagePath = img1Path },
                new Product { Name = "Blue Product", ImagePath = img2Path }
            }
        };

        // Build template document
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Catalog");
        builder.Writeln("<<foreach [p in Products]>>");

        // Table for each product
        Table table = builder.StartTable();

        // Product name cell
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        // Image cell
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // End row and table
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Generate report
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save result
        string resultPath = Path.Combine(outputDir, "ProductCatalogReport.docx");
        doc.Save(resultPath);
    }
}
