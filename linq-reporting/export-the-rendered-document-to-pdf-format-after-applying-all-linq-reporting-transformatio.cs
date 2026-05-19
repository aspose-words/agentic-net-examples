using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using System.Text;

// Ensure code page support for older encodings if needed.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Title = "Product Catalog",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple", Price = 0.99 },
                new Item { Index = 2, Name = "Banana", Price = 0.59 },
                new Item { Index = 3, Name = "Cherry", Price = 2.99 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Export the rendered document to PDF.
        doc.Save("Report.pdf", SaveFormat.Pdf);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title placeholder.
        builder.Writeln("Report: <<[model.Title]>>");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in model.Items]>>");
        // Output each item's details.
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>> - $<<[item.Price]>>");
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model for the report.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Individual item definition.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
