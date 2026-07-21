using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Data model for an order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();

        public Order(string customerName, List<Item> items)
        {
            CustomerName = customerName;
            Items = items;
        }
    }

    // Data model for an item in the order.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }

        public Item(string name, int quantity)
        {
            Name = name;
            Quantity = quantity;
        }
    }

    public static void Main()
    {
        // Step 1: Create a LINQ Reporting template programmatically.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert tags that will be replaced by the reporting engine.
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Items:");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Step 2: Simulate isolated reporting for multiple user requests.
        // Each request gets its own ReportingEngine instance and its own data.
        var userRequests = new[]
        {
            new Order(
                "Alice Johnson",
                new List<Item>
                {
                    new Item("Laptop", 1),
                    new Item("Mouse", 2)
                }),

            new Order(
                "Bob Smith",
                new List<Item>
                {
                    new Item("Desk", 1),
                    new Item("Chair", 4),
                    new Item("Lamp", 2)
                })
        };

        for (int i = 0; i < userRequests.Length; i++)
        {
            // Load a fresh copy of the template for each request.
            var doc = new Document(templatePath);

            // Create a new ReportingEngine instance – isolated per request.
            var engine = new ReportingEngine();

            // Build the report using the order object as the root data source.
            // The template references the root object with the name "order".
            engine.BuildReport(doc, userRequests[i], "order");

            // Save the generated report.
            string outputPath = $"Report_User{i + 1}.docx";
            doc.Save(outputPath);
        }
    }
}
