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
        // Ensure a simple PNG image exists on disk.
        const string imageFileName = "sample.png";
        if (!File.Exists(imageFileName))
        {
            // 1x1 pixel transparent PNG (base64 encoded).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9Yc9Zc8AAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imageFileName, imageBytes);
        }

        // Prepare sample data.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Product A", ImagePath = imageFileName },
                new Product { Name = "Product B", ImagePath = imageFileName }
            }
        };

        // Create the LINQ Reporting template.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [p in Products]>>");
        Table table = builder.StartTable();

        // First cell: product name.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        // Second cell: image inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
