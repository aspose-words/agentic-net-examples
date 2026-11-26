using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingExample
{
    // Data model for an order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
        public decimal Amount { get; set; }
    }

    // Wrapper model passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Sample data: a list of orders with various dates.
            var allOrders = new List<Order>
            {
                new Order { CustomerName = "Alice", OrderDate = DateTime.Now.AddDays(-5), Amount = 120.50m },
                new Order { CustomerName = "Bob",   OrderDate = DateTime.Now.AddDays(-20), Amount = 75.00m },
                new Order { CustomerName = "Carol", OrderDate = DateTime.Now.AddMonths(-2), Amount = 200.00m },
                new Order { CustomerName = "Dave",  OrderDate = DateTime.Now.AddDays(-10), Amount = 50.25m }
            };

            // Filter orders placed within the last month using a lambda expression.
            var recentOrders = allOrders
                .Where(o => o.OrderDate >= DateTime.Now.AddMonths(-1))
                .ToList();

            // Prepare the model for the report.
            var model = new ReportModel { Orders = recentOrders };

            // Create a simple template document with LINQ Reporting tags.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders placed in the last month:");
            builder.Writeln("<<foreach [order in model.Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Date: <<[order.OrderDate]>>");
            builder.Writeln("Amount: $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
