using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model with a collection of orders.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        const string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple report that iterates over a collection named "orders".
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>> (Qty: <<[item.Quantity]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        var orders = new List<Order>
        {
            new Order
            {
                CustomerName = "Alice Johnson",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 3 },
                    new Item { Name = "Banana", Quantity = 5 }
                }
            },
            new Order
            {
                CustomerName = "Bob Smith",
                Items = new List<Item>
                {
                    new Item { Name = "Orange", Quantity = 2 },
                    new Item { Name = "Grapes", Quantity = 1 }
                }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the overload that specifies the collection name.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name of the data source as referenced in the template tags.
        bool success = engine.BuildReport(doc, orders, "orders");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        Console.WriteLine(success
            ? $"Report generated successfully: {reportPath}"
            : "Report generation failed.");
    }
}
