using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for a line item.
    public class LineItem
    {
        public string Product { get; set; } = "";
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    // Data model for an order containing line items.
    public class Order
    {
        public int OrderId { get; set; }
        public string CustomerName { get; set; } = "";
        public List<LineItem> LineItems { get; set; } = new();
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document with nested LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Orders Report ===");
            builder.Writeln();

            // Master band: iterate over orders.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Order ID: <<[order.OrderId]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln();

            // Detail band: iterate over line items of the current order.
            builder.Writeln("<<foreach [item in order.LineItems]>>");
            builder.Writeln("- <<[item.Product]>>  Qty: <<[item.Quantity]>>  Price: $<<[item.Price]>>");
            builder.Writeln("<</foreach>>"); // End of line items foreach.

            builder.Writeln(); // Blank line between orders.
            builder.Writeln("<</foreach>>"); // End of orders foreach.

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data (master-detail relationship).
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new Order
                    {
                        OrderId = 1001,
                        CustomerName = "Alice Johnson",
                        LineItems = new List<LineItem>
                        {
                            new LineItem { Product = "Laptop", Quantity = 1, Price = 1299.99m },
                            new LineItem { Product = "Mouse", Quantity = 2, Price = 25.50m }
                        }
                    },
                    new Order
                    {
                        OrderId = 1002,
                        CustomerName = "Bob Smith",
                        LineItems = new List<LineItem>
                        {
                            new LineItem { Product = "Desk Chair", Quantity = 1, Price = 199.00m },
                            new LineItem { Product = "Monitor", Quantity = 2, Price = 299.99m },
                            new LineItem { Product = "Keyboard", Quantity = 1, Price = 49.99m }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using LINQ Reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            bool success = engine.BuildReport(reportDoc, model, "model");

            // Optionally, you could handle the success flag if InlineErrorMessages were used.
            // For this example we simply save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
