using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Drawing; // Needed for Shape and ShapeType

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample image file that will be used by the data model.
        string sampleImagePath = Path.Combine(outputDir, "SampleImage.png");
        CreateSampleImage(sampleImagePath);

        // 2. Build the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        BuildTemplate(templatePath);

        // 3. Prepare the data model – use image paths directly.
        ReportModel model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Product A", ImagePath = sampleImagePath },
                new Product { Name = "Product B", ImagePath = sampleImagePath }
            }
        };

        // 4. Load the template and run the reporting engine.
        Document templateDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(templateDoc, model, "model");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        templateDoc.Save(reportPath);
    }

    // Creates a simple 1x1 PNG image from a Base64 string.
    private static void CreateSampleImage(string filePath)
    {
        // Base64 for a 1x1 transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/6XcKAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(filePath, pngBytes);
    }

    // Constructs a template that contains a foreach loop over Products.
    // Each iteration creates a table with the product name and its image inside a textbox.
    private static void BuildTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the foreach block.
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table for each product.
        Table table = builder.StartTable();

        // First cell – product name.
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");

        // Second cell – image inside a textbox.
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 120, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Use the image path directly; the engine will load the image from the file system.
        builder.Write("<<image [p.ImagePath] -fitSize>>");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Root data model passed to the reporting engine.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

// Individual product with an image path.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
}
