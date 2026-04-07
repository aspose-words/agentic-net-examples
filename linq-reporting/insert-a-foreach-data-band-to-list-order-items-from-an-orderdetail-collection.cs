using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model representing an order.
    public class Order
    {
        // Sample order identifier.
        public int OrderId { get; set; } = 0;

        // Name of the customer who placed the order.
        public string CustomerName { get; set; } = string.Empty;

        // Collection of order line items.
        public List<OrderDetail> Items { get; set; } = new();
    }

    // Data model for a single line item in an order.
    public class OrderDetail
    {
        // Sequential index of the item (1‑based for display).
        public int Index { get; set; } = 0;

        // Name of the product.
        public string ProductName { get; set; } = string.Empty;

        // Quantity ordered.
        public int Quantity { get; set; } = 0;

        // Unit price.
        public decimal Price { get; set; } = 0m;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Write static content and the LINQ Reporting tags.
            builder.Writeln("Order Report");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order ID: <<[order.OrderId]>>");
            builder.Writeln(); // Empty line for readability.

            // Begin a foreach data band that iterates over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Each iteration writes a line with the item details.
            builder.Writeln("  <<[item.Index]>>. <<[item.ProductName]>> - Qty: <<[item.Quantity]>> - $<<[item.Price]>>");
            // End of the foreach block.
            builder.Writeln("<</foreach>>");

            // 3. Prepare sample data.
            Order sampleOrder = new Order
            {
                OrderId = 1001,
                CustomerName = "John Doe",
                Items = new List<OrderDetail>
                {
                    new OrderDetail { Index = 1, ProductName = "Apple", Quantity = 5, Price = 0.60m },
                    new OrderDetail { Index = 2, ProductName = "Banana", Quantity = 3, Price = 0.40m },
                    new OrderDetail { Index = 3, ProductName = "Cherry", Quantity = 10, Price = 0.15m }
                }
            };

            // 4. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "order".
            engine.BuildReport(doc, sampleOrder, "order");

            // 5. Save the generated report.
            doc.Save("OrderReport.docx");
        }
    }
}
