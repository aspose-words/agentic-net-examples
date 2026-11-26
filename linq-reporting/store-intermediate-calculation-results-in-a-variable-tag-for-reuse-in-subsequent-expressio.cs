using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

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
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Write item details.
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("Quantity: <<[item.Quantity]>>  Unit Price: <<[item.UnitPrice]>>");
        // Compute total price directly in the expression.
        builder.Writeln("Total Price: <<[item.Quantity * item.UnitPrice]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
