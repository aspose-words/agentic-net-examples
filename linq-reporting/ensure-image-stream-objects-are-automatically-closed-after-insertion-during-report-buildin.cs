using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public Stream ImageStream { get; set; } = null!;
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a simple template with a table that contains an image placeholder inside a textbox.
        var templatePath = "template.docx";
        var outputPath = "output.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");

        // Table header
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageStream] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Prepare sample image data (a 1x1 pixel PNG).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YVY6V8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Build the data model.
        var model = new ReportModel();
        for (int i = 1; i <= 3; i++)
        {
            var stream = new MemoryStream(imageBytes);
            // Reset position before the engine reads it.
            stream.Position = 0;

            model.Products.Add(new Product
            {
                Name = $"Product {i}",
                ImageStream = stream
            });
        }

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");
        template.Save(outputPath);

        // Verify that the streams have been closed by the engine.
        bool allClosed = model.Products.All(p =>
        {
            try
            {
                // Attempt to read; should throw if the stream is disposed.
                p.ImageStream.ReadByte();
                return false; // If no exception, stream is still open.
            }
            catch (ObjectDisposedException)
            {
                return true; // Expected outcome.
            }
            catch
            {
                return false;
            }
        });

        Console.WriteLine(allClosed
            ? "All image streams were automatically closed after insertion."
            : "Some image streams remain open.");

        // Dispose any remaining streams just in case.
        foreach (var product in model.Products)
        {
            product.ImageStream?.Dispose();
        }
    }
}
