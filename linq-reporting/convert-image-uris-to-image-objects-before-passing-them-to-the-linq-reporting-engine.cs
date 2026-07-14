using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample image byte arrays (red and blue 1x1 PNG).
        byte[] redPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        byte[] bluePng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/6V6VAAAAAElFTkSuQmCC");

        // Build data model – use image byte arrays for the reporting engine.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Red Square", ImageData = redPng },
                new Product { Name = "Blue Square", ImageData = bluePng }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: Name and Image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Image must be placed inside a textbox.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageData] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source.
        engine.BuildReport(reportDoc, model);

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);
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
    // The reporting engine accepts a byte array containing image data.
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
