using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a tiny PNG image (1x1 pixel) and save it locally.
        string imagePath = Path.Combine(workDir, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XKf8AAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Build the data model.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Product A", ImagePath = imagePath },
                new Product { Name = "Product B", ImagePath = imagePath },
                new Product { Name = "Product C", ImagePath = imagePath }
            }
        };

        // -------------------------
        // Create the LINQ Reporting template.
        // -------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Start a table with two columns: Name and Image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row (will be repeated for each product).
        builder.InsertCell();
        // Insert product name.
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox shape to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag with fitSize switch.
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // End the data row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -------------------------
        // Load the template and build the report.
        // -------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
