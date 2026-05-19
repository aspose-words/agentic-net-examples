using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for the Table class

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
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create sample image files (a tiny 1x1 PNG) that will be used in the report.
        // -----------------------------------------------------------------
        string[] imageFiles = { "product1.png", "product2.png" };
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        foreach (var file in imageFiles)
        {
            File.WriteAllBytes(file, pngBytes);
        }

        // -----------------------------------------------------------------
        // Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Catalog");
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

        // Image must be placed inside a textbox.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        var model = new ReportModel
        {
            Products = new()
            {
                new Product { Name = "Product 1", ImagePath = Path.GetFullPath("product1.png") },
                new Product { Name = "Product 2", ImagePath = Path.GetFullPath("product2.png") }
            }
        };

        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
