using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Product
{
    public string Name { get; set; } = "";
    public Stream ImageStream { get; set; } = Stream.Null;
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a simple 1x1 red PNG image as a byte array.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X" +
            "6V8AAAAASUVORK5CYII=");

        // -----------------------------------------------------------------
        // Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach block that iterates over the Products collection.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: Name and Image.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row – the cells will be filled for each product.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageStream] -fitSize>>");

        // End the table row and the table.
        builder.EndRow();
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Prepare the data model with image streams.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Red Dot", ImageStream = new MemoryStream(pngBytes) },
                new Product { Name = "Red Dot (Copy)", ImageStream = new MemoryStream(pngBytes) }
            }
        };

        // Reset each stream before the report engine consumes it.
        foreach (var product in model.Products)
        {
            if (product.ImageStream.CanSeek)
                product.ImageStream.Position = 0;
        }

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // default options
        };

        // The root object name must match the name used in the template tags ("model").
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        report.Save(outputPath);

        // At this point the ReportingEngine has automatically closed the
        // ImageStream objects after inserting the images, so they are no longer usable.
        // No further action is required.
    }
}
