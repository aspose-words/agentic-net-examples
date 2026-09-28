using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
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
        // Prepare a folder for sample images.
        string imagesDir = Path.Combine(Directory.GetCurrentDirectory(), "Images");
        Directory.CreateDirectory(imagesDir);

        // Create two tiny PNG images from a Base64 string.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lcKcAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(Path.Combine(imagesDir, "product1.png"), pngBytes);
        File.WriteAllBytes(Path.Combine(imagesDir, "product2.png"), pngBytes);

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Catalog");
        builder.Writeln("<<foreach [p in Products]>>");

        // Table that will be repeated for each product.
        Table table = builder.StartTable();

        // First cell – product name.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        // Second cell – image inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product
                {
                    Name = "Product 1",
                    ImagePath = Path.Combine(imagesDir, "product1.png")
                },
                new Product
                {
                    Name = "Product 2",
                    ImagePath = Path.Combine(imagesDir, "product2.png")
                }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
