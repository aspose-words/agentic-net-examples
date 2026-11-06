using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 3, UnitPrice = 0.5m },
                new Item { Name = "Banana", Quantity = 5, UnitPrice = 0.3m },
                new Item { Name = "Cherry", Quantity = 10, UnitPrice = 0.2m }
            }
        };

        // Create the template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header.
        builder.Writeln("Item Report");
        builder.Writeln();

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Use the computed property LineTotal.
        builder.Writeln(
            "Item: <<[item.Name]>> | Qty: <<[item.Quantity]>> | Unit: <<[item.UnitPrice]>> | Total: <<[item.LineTotal]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Root model class.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item class used in the report.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }

    // Computed property for line total.
    public decimal LineTotal => Quantity * UnitPrice;
}
