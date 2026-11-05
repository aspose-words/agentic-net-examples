using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple DOCX template with LINQ Reporting tags.
        var templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new()
                {
                    Name = "Item One",
                    Description = "First sample item.",
                    ImageData = SamplePngBytes()
                },
                new()
                {
                    Name = "Item Two",
                    Description = "Second sample item.",
                    ImageData = SamplePngBytes()
                }
            }
        };

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report as PDF.
        var outputPath = "report.pdf";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Sample LINQ Reporting PDF");
        builder.Writeln();

        // Begin foreach over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Begin table with header
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Description");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row for each item
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");

        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");

        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [item.ImageData] -fitSize>>");

        builder.EndRow();

        // End table
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(path);
    }

    // Returns a byte array of a 1x1 pixel PNG (red color).
    private static byte[] SamplePngBytes()
    {
        return new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
            0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
