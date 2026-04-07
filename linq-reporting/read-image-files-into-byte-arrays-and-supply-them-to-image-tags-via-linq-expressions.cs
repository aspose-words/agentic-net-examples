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
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a tiny PNG image (1x1 pixel, transparent) from a Base64 string
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XG6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Read the image file into a byte array (simulating external image source)
        byte[] imageData = File.ReadAllBytes(imagePath);

        // Prepare sample data
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item 1", ImageData = imageData },
                new Item { Name = "Item 2", ImageData = imageData },
                new Item { Name = "Item 3", ImageData = imageData }
            }
        };

        // Create a Word document template programmatically
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("LINQ Reporting – Image Example");
        builder.Writeln();

        // Begin foreach block
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with two columns: Name and Image
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row (repeated for each item)
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();

        // Insert a textbox to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [item.ImageData] -fitSize>>");

        // Return to the main document flow for the next cell/row
        builder.MoveTo(table.LastRow.LastCell.LastParagraph);

        builder.EndRow();
        builder.EndTable();

        // End foreach block
        builder.Writeln("<</foreach>>");

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated document
        string outputPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine(success
            ? $"Report generated successfully at: {outputPath}"
            : "Report generation failed.");
    }
}

// Root data model
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model containing a name and image data
public class Item
{
    public string Name { get; set; } = string.Empty;
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
