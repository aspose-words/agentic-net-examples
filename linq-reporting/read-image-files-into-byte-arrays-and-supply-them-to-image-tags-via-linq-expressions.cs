using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Product
{
    public string Name { get; set; } = "";
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image from a Base64 string and save it
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Prepare data model
        ReportModel model = new();
        model.Products.Add(new Product
        {
            Name = "Sample Product",
            ImageData = File.ReadAllBytes(imagePath)
        });

        // Create template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Products
        builder.Writeln("<<foreach [p in Products]>>");

        // Create a table with two columns: Name and Image
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();

        // Insert a textbox shape to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 150, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [p.ImageData] -fitSize>>");

        // End the data row
        builder.EndRow();
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the final report
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);
    }
}
