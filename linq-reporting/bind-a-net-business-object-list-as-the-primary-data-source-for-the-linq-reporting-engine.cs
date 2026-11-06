using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingExample
{
    // Business object representing an order.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public DateTime OrderDate { get; set; }
        public List<OrderItem> Items { get; set; } = new();
    }

    // Business object representing an item within an order.
    public class OrderItem
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Step 1: Create the template document programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders Report");
            builder.Writeln();

            // Begin iterating over the collection of orders.
            builder.Writeln("<<foreach [order in orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Date: <<[order.OrderDate]>>");
            builder.Writeln();

            // Table header for order items.
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>> x <<[item.Quantity]>> @ $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Step 2: Prepare sample data (a list of orders).
            List<Order> orders = new()
            {
                new()
                {
                    CustomerName = "John Doe",
                    OrderDate = DateTime.Today,
                    Items = new()
                    {
                        new() { Name = "Widget", Quantity = 2, Price = 9.99m },
                        new() { Name = "Gadget", Quantity = 1, Price = 19.99m }
                    }
                },
                new()
                {
                    CustomerName = "Jane Smith",
                    OrderDate = DateTime.Today.AddDays(-1),
                    Items = new()
                    {
                        new() { Name = "Thingamajig", Quantity = 5, Price = 4.50m }
                    }
                }
            };

            // Step 3: Load the template and build the report.
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, orders, "orders");

            // Step 4: Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
