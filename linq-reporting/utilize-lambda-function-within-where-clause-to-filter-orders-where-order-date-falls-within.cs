using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public DateTime OrderDate { get; set; }
    }

    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample orders
            var allOrders = new List<Order>
            {
                new Order { CustomerName = "Alice", OrderDate = DateTime.Now.AddDays(-5) },
                new Order { CustomerName = "Bob",   OrderDate = DateTime.Now.AddDays(-20) },
                new Order { CustomerName = "Carol", OrderDate = DateTime.Now.AddMonths(-2) } // older than a month
            };

            // Filter orders placed within the last month using a lambda in Where clause
            var recentOrders = allOrders
                .Where(o => o.OrderDate >= DateTime.Now.AddMonths(-1))
                .ToList();

            // Wrap filtered orders in the root model
            var model = new ReportModel { Orders = recentOrders };

            // Create a template document programmatically
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("Orders placed within the last month:");
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>> | Date: <<[order.OrderDate]>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // Load the template
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
