using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Product
{
    public string Name { get; set; } = string.Empty;
    public MemoryStream ImageStream { get; set; } = new();
}

public class ReportModel : IDisposable
{
    public List<Product> Products { get; set; } = new();

    public void Dispose()
    {
        foreach (var p in Products)
        {
            p.ImageStream?.Dispose();
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        Directory.CreateDirectory("output");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin the foreach block for products.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table for each product row.
        Table table = builder.StartTable();

        // Header row (only once, before the loop repeats).
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row: product name.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        // Data row: image inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageStream] -fitSize>>");

        // End of the data row.
        builder.EndRow();

        // End the table for this iteration.
        builder.EndTable();

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save("template.docx");

        // -----------------------------------------------------------------
        // 2. Load the template back (as required by the workflow).
        // -----------------------------------------------------------------
        var doc = new Document("template.docx");

        // -----------------------------------------------------------------
        // 3. Prepare sample image data (a tiny red dot PNG).
        // -----------------------------------------------------------------
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9WcKXKcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // -----------------------------------------------------------------
        // 4. Create sample products with image streams.
        // -----------------------------------------------------------------
        var products = new List<Product>
        {
            new Product { Name = "Product A", ImageStream = new MemoryStream(imageBytes, writable: false) },
            new Product { Name = "Product B", ImageStream = new MemoryStream(imageBytes, writable: false) },
            new Product { Name = "Product C", ImageStream = new MemoryStream(imageBytes, writable: false) }
        };

        // Reset each stream position before the report engine reads them.
        foreach (var p in products)
        {
            if (p.ImageStream != null)
                p.ImageStream.Position = 0;
        }

        // -----------------------------------------------------------------
        // 5. Build the report. Streams will be disposed automatically when the model is disposed.
        // -----------------------------------------------------------------
        using (var model = new ReportModel { Products = products })
        {
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");
            doc.Save("output/report.docx");
        }

        // At this point all image streams have been closed automatically.
    }
}
