using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple order data model.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data with various dates.
            List<Order> allOrders = new()
            {
                new Order { CustomerName = "Alice",   OrderDate = DateTime.Today.AddDays(-5) },
                new Order { CustomerName = "Bob",     OrderDate = DateTime.Today.AddDays(-20) },
                new Order { CustomerName = "Charlie", OrderDate = DateTime.Today.AddMonths(-2) },
                new Order { CustomerName = "Diana",   OrderDate = DateTime.Today.AddDays(-1) },
                new Order { CustomerName = "Eve",     OrderDate = DateTime.Today.AddMonths(-1).AddDays(-1) }
            };

            // Use a lambda expression in a Where clause to keep only orders from the last month.
            DateTime today = DateTime.Today;
            DateTime monthAgo = today.AddMonths(-1);
            List<Order> recentOrders = allOrders
                .Where(o => o.OrderDate >= monthAgo && o.OrderDate <= today)
                .ToList();

            // Wrap the filtered collection in the model.
            ReportModel model = new()
            {
                Orders = recentOrders
            };

            // -----------------------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders placed within the last month:");
            // Note the use of the root name "model" as required by the engine.
            builder.Writeln("<<foreach [order in model.Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>> | Date: <<[order.OrderDate]>>");
            builder.Writeln("<</foreach>>");

            // Save the template before building the report (lifecycle rule).
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
