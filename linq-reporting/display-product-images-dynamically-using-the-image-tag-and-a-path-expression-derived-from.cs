using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample image.
        string imageFileName = "product.png";
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), imageFileName);
        CreateSamplePng(imagePath);

        // Prepare data model.
        ReportModel model = new()
        {
            Products = new()
            {
                new Product { Name = "Sample Product", ImagePath = imagePath }
            }
        };

        // Create template document.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load template and build report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }

    private static void CreateSamplePng(string path)
    {
        // 1x1 pixel red PNG.
        const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6ZcAAAAASUVORK5CYII=";
        byte[] bytes = Convert.FromBase64String(base64);
        File.WriteAllBytes(path, bytes);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Begin foreach over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two cells: Name and Image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox shape to hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 150);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
