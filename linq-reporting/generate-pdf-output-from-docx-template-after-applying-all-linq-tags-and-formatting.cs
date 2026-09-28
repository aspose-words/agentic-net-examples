using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample image (1x1 PNG) as a byte array.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        // Optional: write the image to disk for reference.
        File.WriteAllBytes("sample.png", pngBytes);

        // Build sample data model.
        ReportModel model = new()
        {
            Title = "Sales Report",
            CustomerName = "Acme Corp",
            DescriptionHtml = "<b>Quarterly performance summary.</b>",
            Items = new()
            {
                new Item { Index = 1, Name = "Widget", Price = 75.00m },
                new Item { Index = 2, Name = "Gadget", Price = 150.00m },
                new Item { Index = 3, Name = "Doohickey", Price = 45.00m }
            },
            Images = new()
            {
                new ImageData { Data = pngBytes },
                new ImageData { Data = pngBytes }
            }
        };

        // -----------------------------------------------------------------
        // Create DOCX template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new();
        DocumentBuilder builder = new(template);

        // Simple text fields.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("<<[model.DescriptionHtml] -html>>");
        builder.Writeln();

        // Items table – repeated rows.
        builder.Writeln("<<foreach [item in Items]>>");
        Table itemsTable = builder.StartTable();

        // Index cell.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");

        // Name cell.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");

        // Price cell with conditional background.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Price > 100]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Price]>> <</backColor>><</if>>" +
            "<<if [item.Price <= 100]>>" +
            "<<[item.Price]>> <</if>>");

        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Images section – each image inside a textbox within a table cell.
        builder.Writeln("<<foreach [img in Images]>>");
        Table imgTable = builder.StartTable();
        builder.InsertCell();

        Shape txtBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(txtBox.FirstParagraph);
        builder.Writeln("<<image [img.Data] -fitSize>>");

        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the final report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new()
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Ensure output directory exists.
        Directory.CreateDirectory("output");

        // Save the final report as PDF.
        string pdfPath = Path.Combine("output", "Report.pdf");
        reportDoc.Save(pdfPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public string Title { get; set; } = "";
    public string CustomerName { get; set; } = "";
    public string DescriptionHtml { get; set; } = "";
    public List<Item> Items { get; set; } = new();
    public List<ImageData> Images { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}

public class ImageData
{
    public byte[] Data { get; set; } = Array.Empty<byte>();
}
