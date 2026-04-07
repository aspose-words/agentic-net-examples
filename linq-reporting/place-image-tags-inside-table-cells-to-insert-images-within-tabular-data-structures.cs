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
        // Prepare a tiny red PNG image file.
        const string imageFileName = "sample.png";
        if (!File.Exists(imageFileName))
        {
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9W6V6WcAAAAASUVORK5CYII=";
            File.WriteAllBytes(imageFileName, Convert.FromBase64String(base64Png));
        }

        // Create the data source.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  ImagePath = imageFileName },
                new Product { Name = "Banana", ImagePath = imageFileName },
                new Product { Name = "Cherry", ImagePath = imageFileName }
            }
        };

        // Build the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin the foreach loop.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: Name and Image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Product Name");
        builder.InsertCell();
        builder.Write("Image");
        builder.EndRow();

        // Data row (repeated for each product).
        builder.InsertCell();
        builder.Write("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // End the data row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for clarity).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "ReportWithImages.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
