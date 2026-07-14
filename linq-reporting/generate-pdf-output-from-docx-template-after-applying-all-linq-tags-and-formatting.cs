using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple PNG image (1x1 transparent pixel) and save it locally.
        const string imageFileName = "sample.png";
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        File.WriteAllBytes(imageFileName, Convert.FromBase64String(base64Png));

        // Build the DOCX template programmatically.
        const string templateFileName = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title tag.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // HTML snippet tag (rendered as HTML).
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Table of items using a foreach loop.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln(" - <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Image tag placed inside a textbox container.
        builder.Writeln("Image:");
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // Load the template for reporting.
        var doc = new Document(templateFileName);

        // Prepare sample data model.
        var model = new ReportModel
        {
            Title = "Sample LINQ Reporting",
            HtmlSnippet = "<b>Bold HTML Content</b> and <i>italic text</i>",
            ImagePath = imageFileName,
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.20m },
                new Item { Name = "Banana", Price = 0.80m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the populated document as PDF.
        const string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);
    }
}

// Data model classes must be public with public properties.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string HtmlSnippet { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
