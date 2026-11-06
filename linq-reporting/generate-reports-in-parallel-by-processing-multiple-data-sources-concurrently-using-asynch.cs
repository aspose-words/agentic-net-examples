using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingParallel
{
    // Data model for an order.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Data model for an item within an order.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class Program
    {
        // Entry point.
        public static async Task Main(string[] args)
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create the template document programmatically.
            string templatePath = "ReportTemplate.docx";
            CreateTemplate(templatePath);

            // Prepare sample data sources.
            List<Order> orders = GetSampleOrders();

            // Generate reports in parallel.
            List<Task> tasks = new();
            for (int i = 0; i < orders.Count; i++)
            {
                int index = i; // Capture loop variable.
                tasks.Add(GenerateReportAsync(orders[index], index, templatePath));
            }

            await Task.WhenAll(tasks);
        }

        // Creates a simple LINQ Reporting template and saves it to disk.
        private static void CreateTemplate(string path)
        {
            Document doc = new();
            DocumentBuilder builder = new(doc);

            // Customer name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Table header inside foreach.
            builder.Writeln("<<foreach [item in order.Items]>>");
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Table rows.
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }

        // Generates a report for a single order asynchronously.
        private static async Task GenerateReportAsync(Order order, int orderIndex, string templatePath)
        {
            await Task.Run(() =>
            {
                // Load the template.
                Document doc = new(templatePath);

                // Build the report.
                ReportingEngine engine = new();
                engine.BuildReport(doc, order, "order");

                // Save the generated report.
                string outputPath = $"Report_{orderIndex + 1}.docx";
                doc.Save(outputPath);
                Console.WriteLine($"Report generated: {outputPath}");
            });
        }

        // Returns a list of sample orders.
        private static List<Order> GetSampleOrders()
        {
            return new()
            {
                new Order
                {
                    CustomerName = "Alice Johnson",
                    Items = new()
                    {
                        new Item { Index = 1, Name = "Widget A" },
                        new Item { Index = 2, Name = "Widget B" }
                    }
                },
                new Order
                {
                    CustomerName = "Bob Smith",
                    Items = new()
                    {
                        new Item { Index = 1, Name = "Gadget X" },
                        new Item { Index = 2, Name = "Gadget Y" },
                        new Item { Index = 3, Name = "Gadget Z" }
                    }
                },
                new Order
                {
                    CustomerName = "Carol Davis",
                    Items = new()
                    {
                        new Item { Index = 1, Name = "Thingamajig" }
                    }
                }
            };
        }
    }
}
