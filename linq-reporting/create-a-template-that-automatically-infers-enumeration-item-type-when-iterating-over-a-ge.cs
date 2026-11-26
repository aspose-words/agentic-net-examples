using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class Order
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Order Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
