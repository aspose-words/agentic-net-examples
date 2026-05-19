using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Price = 1.23m },
                new Item { Name = "Banana", Price = 0.99m },
                new Item { Name = "Cherry", Price = 2.50m }
            }
        };

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        CreateTemplate(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Output product name.
        builder.Writeln("Product: <<[item.Name]>>");

        // Output price using a custom currency format.
        // The format string is placed inside double quotes to be parsed correctly.
        builder.Writeln("Price: <<[item.Price]:\"$#,0.00\">>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
