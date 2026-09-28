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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // Create sample business object list as data source.
        // -------------------------------------------------
        List<Order> orders = new()
        {
            new Order
            {
                CustomerName = "Alice",
                Items = new()
                {
                    new Item { Index = 1, Name = "Item A" },
                    new Item { Index = 2, Name = "Item B" }
                }
            },
            new Order
            {
                CustomerName = "Bob",
                Items = new()
                {
                    new Item { Index = 1, Name = "Item X" }
                }
            }
        };

        // -------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, orders, "orders");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}

// -------------------------------------------------
// Public data model classes.
// -------------------------------------------------
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
