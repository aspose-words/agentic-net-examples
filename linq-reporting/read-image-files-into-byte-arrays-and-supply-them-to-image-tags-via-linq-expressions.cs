using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public byte[] Image { get; set; } = Array.Empty<byte>();
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
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create two 1x1 PNG images from Base64 strings.
        string redPixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        string greenPixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lZcVAAAAAElFTkSuQmCC";

        string redPath = Path.Combine(workDir, "red.png");
        string greenPath = Path.Combine(workDir, "green.png");

        File.WriteAllBytes(redPath, Convert.FromBase64String(redPixelBase64));
        File.WriteAllBytes(greenPath, Convert.FromBase64String(greenPixelBase64));

        // Build the data model.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Red Pixel", Image = File.ReadAllBytes(redPath) },
                new Product { Name = "Green Pixel", Image = File.ReadAllBytes(greenPath) }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin the foreach block.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table row for each product.
        builder.StartTable();

        // First cell – product name.
        builder.InsertCell();
        builder.Write("Name: <<[p.Name]>>");

        // Second cell – image inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.Image] -fitSize>>");

        // End the table row.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(workDir, "report.docx");
        reportDoc.Save(reportPath);
    }
}
