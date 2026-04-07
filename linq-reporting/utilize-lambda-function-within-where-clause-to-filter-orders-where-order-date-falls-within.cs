using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for an order.
    public class Order
    {
        public int Id { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
        public decimal Amount { get; set; }
    }

    // Wrapper class that will be passed to the reporting engine.
    public class ReportData
    {
        public List<Order> Orders { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample orders with various dates.
            List<Order> allOrders = new()
            {
                new Order { Id = 1, CustomerName = "Alice",   OrderDate = DateTime.Today.AddDays(-5),  Amount = 120.50m },
                new Order { Id = 2, CustomerName = "Bob",     OrderDate = DateTime.Today.AddDays(-20), Amount = 75.00m },
                new Order { Id = 3, CustomerName = "Charlie", OrderDate = DateTime.Today.AddMonths(-2), Amount = 200.00m },
                new Order { Id = 4, CustomerName = "Diana",   OrderDate = DateTime.Today.AddDays(-15), Amount = 45.30m }
            };

            // Use a lambda expression to keep only orders from the last month.
            DateTime start = DateTime.Today.AddMonths(-1);
            DateTime end   = DateTime.Today;
            List<Order> recentOrders = allOrders
                .Where(o => o.OrderDate >= start && o.OrderDate < end)
                .ToList();

            // Wrap the filtered collection.
            ReportData data = new() { Orders = recentOrders };

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = "ReportTemplate.docx";
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Recent Orders (last month):");
            // The root object will be referenced as 'model' in the template.
            builder.Writeln("<<foreach [order in model.Orders]>>");
            builder.Writeln(
                "Order ID: <<[order.Id]>>, " +
                "Customer: <<[order.CustomerName]>>, " +
                "Date: <<[order.OrderDate]>>, " +
                "Amount: $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport must be called after the template is fully prepared.
            engine.BuildReport(reportDoc, data, "model");

            // Save the generated report.
            string outputPath = "ReportOutput.docx";
            reportDoc.Save(outputPath);

            // Inform the user where the files are located (no interactive input required).
            Console.WriteLine($"Template saved to: {Path.GetFullPath(templatePath)}");
            Console.WriteLine($"Report generated at: {Path.GetFullPath(outputPath)}");
        }
    }
}
