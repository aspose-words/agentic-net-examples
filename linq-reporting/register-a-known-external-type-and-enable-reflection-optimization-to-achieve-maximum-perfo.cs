using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // External static helper that will be registered as a known type.
    public static class Formatter
    {
        public static string Upper(string value) => value?.ToUpperInvariant() ?? string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Enable reflection optimization for maximum performance.
            ReportingEngine.UseReflectionOptimization = true;

            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[Formatter.Upper(item.Name)]>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            Order order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "apple" },
                    new Item { Index = 2, Name = "banana" },
                    new Item { Index = 3, Name = "cherry" }
                }
            };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Formatter));

            // Build the report using the root object name "order".
            engine.BuildReport(template, order, "order");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
