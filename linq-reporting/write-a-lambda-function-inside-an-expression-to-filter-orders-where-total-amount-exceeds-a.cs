using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Order data model
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = string.Empty;
        public decimal Total { get; set; } = 0m;
    }

    // Root model passed to the reporting engine
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
        public decimal Threshold { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice", Total = 75.50m },
                new Order { Id = 2, CustomerName = "Bob",   Total = 150.00m },
                new Order { Id = 3, CustomerName = "Carol", Total = 200.25m },
                new Order { Id = 4, CustomerName = "Dave",  Total = 45.00m }
            };

            var model = new ReportModel
            {
                Orders = orders,
                Threshold = 100m // Filter threshold
            };

            // Create the template document programmatically
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Orders with total greater than the threshold:");
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.Total > model.Threshold)]>>");
            builder.Writeln("Order ID: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Total: <<[order.Total]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template and build the report
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(loadedDoc, model, "model");

            // Save the generated report
            var outputPath = "report.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
