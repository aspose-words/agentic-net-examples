using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words on .NET Core)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple", Price = 0.5m, Quantity = 4 },
                new Item { Name = "Banana", Price = 0.3m, Quantity = 6 },
                new Item { Name = "Cherry", Price = 1.2m, Quantity = 10 }
            }
        };

        // Create template document
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("Quantity: <<[item.Quantity]>>");
        builder.Writeln("Price per unit: <<[item.Price]>>");
        // Use a calculated property instead of an unsupported let tag
        builder.Writeln("Line total: <<[item.LineTotal]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Build the report
        Document report = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(report, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Quantity { get; set; }

    // Calculated property used in the template
    public decimal LineTotal => Price * Quantity;
}
