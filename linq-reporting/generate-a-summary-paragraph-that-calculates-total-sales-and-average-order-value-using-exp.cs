using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSummary
{
    // Root data model for the report.
    public class ReportModel
    {
        // Collection of orders.
        public List<Order> Orders { get; set; } = new();

        // Total sales calculated from the orders.
        public decimal TotalSales => Orders.Sum(o => o.Total);

        // Average order value (total sales divided by number of orders).
        public decimal AverageOrderValue => Orders.Count > 0 ? TotalSales / Orders.Count : 0;
    }

    // Simple order class.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public decimal UnitPrice { get; set; }
        public int Quantity { get; set; }

        // Total amount for this order.
        public decimal Total => UnitPrice * Quantity;
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new Order { CustomerName = "Alice", UnitPrice = 120.50m, Quantity = 2 },
                    new Order { CustomerName = "Bob",   UnitPrice =  75.00m, Quantity = 1 },
                    new Order { CustomerName = "Carol", UnitPrice =  45.99m, Quantity = 5 }
                }
            };

            // Create a new blank document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a title.
            builder.Writeln("Sales Summary");
            builder.Writeln();

            // Insert the summary paragraph with expression tags.
            // The tags reference the root object named "model".
            builder.Writeln("Total Sales: <<[model.TotalSales]>>");
            builder.Writeln("Average Order Value: <<[model.AverageOrderValue]>>");

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "SalesSummaryReport.docx";
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
