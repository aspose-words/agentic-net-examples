using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model classes must be public with public properties.
    public class Order
    {
        // Initialize non‑nullable reference types to avoid warnings.
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item: <<[item.Name]>> - Qty: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            Order order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 3 },
                    new Item { Name = "Banana", Quantity = 5 },
                    new Item { Name = "Orange", Quantity = 2 }
                }
            };

            // 3. Enable reflection optimization for large collections.
            ReportingEngine.UseReflectionOptimization = true;

            // 4. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, order, "order");

            // 5. Save the generated report.
            template.Save("Report_ReflectionOptimization.docx");
        }
    }
}
