using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template back before building the report.
        Document templateDoc = new Document(templatePath);

        // Prepare sample data: a list of products with Base64‑encoded images.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product
                {
                    Name = "Red Dot",
                    // 1×1 red PNG (Base64 encoded).
                    ImageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII="
                },
                new Product
                {
                    Name = "Green Dot",
                    // 1×1 green PNG (Base64 encoded).
                    ImageBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXK5wAAAABJRU5ErkJggg=="
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        templateDoc.Save("Report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin foreach loop over Products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Table with two columns: Name and Image.
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

        // Insert a textbox to host the image tag (required by the engine).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageBytes] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Individual product with a Base64 image string.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public string ImageBase64 { get; set; } = string.Empty;

    // Convert the Base64 string to a byte array for the image tag.
    public byte[] ImageBytes => Convert.FromBase64String(ImageBase64);
}
