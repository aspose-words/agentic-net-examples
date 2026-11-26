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
        // -----------------------------------------------------------------
        // Create sample PNG image data (1x1 pixel) and write two files.
        // -----------------------------------------------------------------
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV6ZbUAAAAASUVORK5CYII=");
        string[] imageFiles = { "product1.png", "product2.png" };
        foreach (var file in imageFiles)
        {
            File.WriteAllBytes(file, pngBytes);
        }

        // -----------------------------------------------------------------
        // Prepare the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Product 1", ImagePath = Path.GetFullPath("product1.png") },
                new Product { Name = "Product 2", ImagePath = Path.GetFullPath("product2.png") }
            }
        };

        // -----------------------------------------------------------------
        // Build the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin the foreach block.
        builder.Writeln("<<foreach [p in Products]>>");

        // Start a table for each product.
        Table table = builder.StartTable();

        // Header row (will be repeated for each product – acceptable for demo).
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox shape that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk and reload it for reporting.
        const string templatePath = "template.docx";
        template.Save(templatePath);
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
}
