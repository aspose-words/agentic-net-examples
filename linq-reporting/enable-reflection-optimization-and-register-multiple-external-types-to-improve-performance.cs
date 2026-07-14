using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model classes
    public class Order
    {
        public int Id { get; set; } = 0;
        public Customer Customer { get; set; } = new();
        public List<Item> Items { get; set; } = new();
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    // External static helper types that will be used inside the template
    public static class Helper
    {
        public static string FormatPrice(decimal price) => $"${price:F2}";
    }

    public static class MathHelper
    {
        public static int Double(int value) => value * 2;
    }

    public class Program
    {
        public static void Main()
        {
            // Enable reflection optimization (static property)
            ReportingEngine.UseReflectionOptimization = true;

            // Create a simple template document in memory
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Order Report");
            builder.Writeln("==============");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.Customer.Name]>>");
            builder.Writeln("Email: <<[order.Customer.Email]>>");
            builder.Writeln();
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- Name: <<[item.Name]>>");
            builder.Writeln("  Price: <<[Helper.FormatPrice(item.Price)]>>");
            builder.Writeln("  Quantity (x2): <<[MathHelper.Double(item.Quantity)]>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample hierarchical data
            Order sampleOrder = new()
            {
                Id = 12345,
                Customer = new Customer
                {
                    Name = "John Doe",
                    Email = "john.doe@example.com"
                },
                Items = new List<Item>
                {
                    new Item { Name = "Widget", Price = 19.99m, Quantity = 2 },
                    new Item { Name = "Gadget", Price = 34.50m, Quantity = 1 },
                    new Item { Name = "Doohickey", Price = 5.75m, Quantity = 5 }
                }
            };

            // Configure the reporting engine
            ReportingEngine engine = new ReportingEngine();

            // Register external types so that static members can be used in the template
            engine.KnownTypes.Add(typeof(Helper));
            engine.KnownTypes.Add(typeof(MathHelper));

            // Build the report using the root object name "order"
            engine.BuildReport(doc, sampleOrder, "order");

            // Ensure output directory exists
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the generated report
            string outputPath = Path.Combine(outputDir, "OrderReport.docx");
            doc.Save(outputPath);
        }
    }
}
