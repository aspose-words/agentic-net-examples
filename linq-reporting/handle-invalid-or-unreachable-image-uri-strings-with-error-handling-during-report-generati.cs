using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a tiny PNG image file for the valid case.
        // -----------------------------------------------------------------
        string validImagePath = Path.Combine(outputDir, "valid.png");
        // 1x1 transparent PNG (base64 encoded).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(validImagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection named "Products".
        builder.Writeln("<<foreach [p in Products]>>");

        // Write product name.
        builder.Writeln("Product: <<[p.Name]>>");

        // Insert a textbox that will contain the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag – the expression may resolve to a valid file path or an invalid URI.
        builder.Write("<<image [p.ImageUri] -fitSize>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Prepare sample data with one valid and one invalid image URI.
        // -----------------------------------------------------------------
        var products = new List<Product>
        {
            new Product("Valid Image", validImagePath),                     // Local file – valid.
            new Product("Invalid Image", "http://invalid.example.com/img.png") // Unreachable URI – invalid.
        };

        // Wrap the collection in a public model class (required for LINQ Reporting).
        var model = new ReportModel { Products = products };

        // -----------------------------------------------------------------
        // 5. Build the report using InlineErrorMessages to capture errors.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // The root object name must match the name used in the template ("Products").
        bool success = engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(resultPath);

        // Output the success flag – no interactive prompts are used.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Report saved to: {resultPath}");
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class Product
{
    public Product(string name, string imageUri)
    {
        Name = name;
        ImageUri = imageUri;
    }

    public string Name { get; set; }
    public string ImageUri { get; set; }
}

// Wrapper class required because LINQ Reporting does not accept anonymous types.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}
