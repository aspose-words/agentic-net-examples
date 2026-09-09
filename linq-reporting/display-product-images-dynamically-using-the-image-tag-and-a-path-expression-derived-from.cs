using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
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
        // Prepare a working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "work");
        Directory.CreateDirectory(workDir);

        // Create three tiny PNG files (1x1 pixel) from a Base64 string.
        string[] imageNames = { "apple.png", "banana.png", "cherry.png" };
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3V8AAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        foreach (string name in imageNames)
        {
            File.WriteAllBytes(Path.Combine(workDir, name), pngBytes);
        }

        // Build the data model.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  ImagePath = Path.Combine(workDir, "apple.png") },
                new Product { Name = "Banana", ImagePath = Path.Combine(workDir, "banana.png") },
                new Product { Name = "Cherry", ImagePath = Path.Combine(workDir, "cherry.png") }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach loop over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: product name and product image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row (repeated for each product).
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        builder.InsertCell();
        // Insert a textbox that will host the image tag.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with fitSize switch.
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // Finish the row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(workDir, "ProductTemplate.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the final report.
        string outputPath = Path.Combine(workDir, "ProductReport.docx");
        report.Save(outputPath);

        Console.WriteLine("Report generated at: " + outputPath);
    }
}
