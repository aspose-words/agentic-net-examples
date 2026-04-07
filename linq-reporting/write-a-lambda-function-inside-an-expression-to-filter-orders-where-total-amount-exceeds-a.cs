using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLambdaExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of orders.
        public List<Order> Orders { get; set; } = new();

        // Threshold used for filtering.
        public decimal Threshold { get; set; }
    }

    // Simple order class.
    public class Order
    {
        public int Id { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public decimal Total { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice", Total = 750m },
                new Order { Id = 2, CustomerName = "Bob",   Total = 1250m },
                new Order { Id = 3, CustomerName = "Carol", Total = 2000m },
                new Order { Id = 4, CustomerName = "Dave",  Total = 500m }
            };

            var model = new ReportModel
            {
                Orders = orders,
                Threshold = 1000m // Only orders with Total > 1000 will be shown.
            };

            // 2. Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Orders with total amount greater than the threshold:");
            // Lambda expression inside the foreach tag filters the collection.
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.Total > model.Threshold)]>>");
            builder.Writeln("Order ID: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Total: <<[order.Total]>>");
            builder.Writeln("<</foreach>>");

            // 3. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated report.
            doc.Save("FilteredOrdersReport.docx");
        }
    }
}
