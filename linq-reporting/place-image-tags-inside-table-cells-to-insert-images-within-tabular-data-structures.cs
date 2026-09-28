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
        // Create output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create two tiny PNG files to be used as images
        string[] imageFiles =
        {
            Path.Combine(outputDir, "image1.png"),
            Path.Combine(outputDir, "image2.png")
        };
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YVhZ3cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        foreach (var path in imageFiles)
        {
            File.WriteAllBytes(path, pngBytes);
        }

        // Prepare the data model
        ReportModel model = new()
        {
            Products = new()
            {
                new Product { Name = "Product A", ImagePath = imageFiles[0] },
                new Product { Name = "Product B", ImagePath = imageFiles[1] }
            }
        };

        // Build the template document
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document template = new();
        DocumentBuilder builder = new(template);

        // Open foreach block
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table for each product row
        Table table = builder.StartTable();

        // Header row (only once, but placed inside foreach for simplicity)
        builder.InsertCell();
        builder.Writeln("Product Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox shape that will hold the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("<<image [p.ImagePath] -fitSize>>");

        // Finish the row and table
        builder.EndRow();
        builder.EndTable();

        // Close foreach block
        builder.Writeln("<</foreach>>");

        // Save the template
        template.Save(templatePath);

        // Load the template and generate the report
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report
        string reportPath = Path.Combine(outputDir, "Report.docx");
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
    public string Name { get; set; } = "";
    public string ImagePath { get; set; } = "";
}
