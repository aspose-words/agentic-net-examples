using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineIsolationExample
{
    // Sample data model for an order.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Sample data model for an item.
    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
    }

    public class Program
    {
        // Path constants.
        private const string TemplatePath = "Template.docx";
        private const string Report1Path = "Report1.docx";
        private const string Report2Path = "Report2.docx";

        public static void Main()
        {
            // Step 1: Create the LINQ Reporting template programmatically.
            CreateTemplate();

            // Step 2: Process first request with its own ReportingEngine instance.
            var order1 = new Order
            {
                CustomerName = "Alice Johnson",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Price = 1.20m },
                    new Item { Name = "Banana", Price = 0.80m }
                }
            };
            GenerateReport(order1, Report1Path);

            // Step 3: Process second request with a separate ReportingEngine instance.
            var order2 = new Order
            {
                CustomerName = "Bob Smith",
                Items = new List<Item>
                {
                    new Item { Name = "Orange", Price = 1.50m },
                    new Item { Name = "Grapes", Price = 2.30m },
                    new Item { Name = "Mango", Price = 3.00m }
                }
            };
            GenerateReport(order2, Report2Path);
        }

        // Creates a Word document containing LINQ Reporting tags.
        private static void CreateTemplate()
        {
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert static text and data placeholders.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(TemplatePath);
        }

        // Generates a report for a given order using an isolated ReportingEngine.
        private static void GenerateReport(Order order, string outputPath)
        {
            // Load the previously saved template.
            var doc = new Document(TemplatePath);

            // Each request gets its own ReportingEngine instance.
            var engine = new ReportingEngine();

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
