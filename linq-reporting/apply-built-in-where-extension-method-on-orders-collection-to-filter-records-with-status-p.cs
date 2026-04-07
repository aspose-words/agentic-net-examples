using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model representing a single order.
    public class Order
    {
        public string OrderId { get; set; } = "";
        public string CustomerName { get; set; } = "";
        public string Status { get; set; } = "";
    }

    // Wrapper class that holds the collection of orders.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var orders = new List<Order>
            {
                new Order { OrderId = "1001", CustomerName = "Alice",   Status = "Pending" },
                new Order { OrderId = "1002", CustomerName = "Bob",     Status = "Shipped" },
                new Order { OrderId = "1003", CustomerName = "Charlie", Status = "Pending" },
                new Order { OrderId = "1004", CustomerName = "Diana",   Status = "Cancelled" }
            };

            var model = new ReportModel { Orders = orders };

            // Create a template document programmatically.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            CreateTemplate(templatePath);

            // Load the template.
            Document doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(reportPath);
        }

        // Generates a Word document containing LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Orders with status 'Pending':");
            // Apply the built‑in Where extension method to filter the collection.
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.Status == \"Pending\")]>>");
            builder.Writeln("Order ID: <<[order.OrderId]>>, Customer: <<[order.CustomerName]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
