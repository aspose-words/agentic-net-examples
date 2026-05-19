using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingDemo
{
    // Sample data model classes.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = string.Empty;
        public List<OrderItem> Items { get; set; } = new();
        public DateTime OrderDate { get; set; } = DateTime.Now;
    }

    public class OrderItem
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; } = 0;
        public decimal Price { get; set; } = 0m;
    }

    // Wrapper model used as the root data source for the LINQ Reporting engine.
    public class ReportModel
    {
        public Order Order { get; set; } = new();
        // JSON representation of the Order object for debugging.
        public string JsonDebug { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample data.
            // -----------------------------------------------------------------
            var order = new Order
            {
                Id = 1001,
                CustomerName = "John Doe",
                OrderDate = DateTime.Today,
                Items = new List<OrderItem>
                {
                    new OrderItem { Name = "Widget A", Quantity = 3, Price = 9.99m },
                    new OrderItem { Name = "Widget B", Quantity = 1, Price = 19.95m }
                }
            };

            // Serialize the order to JSON (indented for readability).
            string json = JsonConvert.SerializeObject(order, Formatting.Indented);

            // Wrap the data for the reporting engine.
            var model = new ReportModel
            {
                Order = order,
                JsonDebug = json
            };

            // -----------------------------------------------------------------
            // 2. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("=== Order Debug Information ===");
            // Insert a LINQ Reporting tag that will output the JSON string.
            builder.Writeln("<<[model.JsonDebug]>>");

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the resulting document.
            // -----------------------------------------------------------------
            const string outputPath = "OrderReport.docx";
            doc.Save(outputPath);
        }
    }
}
