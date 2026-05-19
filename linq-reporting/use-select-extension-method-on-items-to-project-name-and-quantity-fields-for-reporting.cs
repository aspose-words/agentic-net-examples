using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var items = new List<Item>
        {
            new() { Name = "Apple", Quantity = 10, Price = 0.5m },
            new() { Name = "Banana", Quantity = 5, Price = 0.3m },
            new() { Name = "Orange", Quantity = 8, Price = 0.6m }
        };

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Use Select to project only Name and Quantity fields.
        builder.Writeln("<<foreach [item in Items.Select(i => new { Name = i.Name, Quantity = i.Quantity })]>>");
        builder.Writeln("Product: <<[item.Name]>> | Qty: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Prepare the root data model.
        var model = new ReportModel { Items = items };

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
