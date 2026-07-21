using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;          // Needed for ShapeType
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image (a 1x1 red pixel) for the valid case.
        string validImagePath = Path.Combine(outputDir, "valid.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YV4cVIAAAAASUVORK5CYII=");
        File.WriteAllBytes(validImagePath, pngBytes);

        // Build the data model.
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Valid Image", ImageUri = validImagePath },
                new Product { Name = "Invalid Image", ImageUri = "nonexistent_image.jpg" }
            }
        };

        // Create the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");

        // Insert a textbox that will hold the image.
        var textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag – the expression returns a string (file path or URL).
        builder.Write("<<image [p.ImageUri] -fitSize>>");

        // Return to the main story to write the product name.
        builder.MoveToDocumentEnd();
        builder.Writeln("<<[p.Name]>>");

        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The returned flag indicates whether parsing succeeded.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);

        // Output the success flag (for demonstration; not required by the task).
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {resultPath}");
    }
}

// Root data model.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Simple product class.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public string ImageUri { get; set; } = string.Empty;
}
