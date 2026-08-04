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
        // Sample data with a Base64 encoded 1x1 PNG.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product
                {
                    Name = "Sample Product",
                    ImageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9ZcAAAAASUVORK5CYII="
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln(string.Empty);

        // Begin foreach block.
        builder.Writeln("<<foreach [p in Products]>>");

        // Table with two columns: Name and Image.
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

        // Insert a textbox that will hold the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageBytes] -fitSize>>");

        // Return to the table cell after the textbox.
        builder.MoveTo(table.LastRow.LastCell.LastParagraph);
        builder.EndRow();
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
        reportDoc.Save(outputPath);
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
    public string ImageBase64 { get; set; } = string.Empty;

    // Convert Base64 string to byte array for the image tag.
    public byte[] ImageBytes => Convert.FromBase64String(ImageBase64);
}
