using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Data model for an order.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = "";
        public string Status { get; set; } = "";
        public decimal Amount { get; set; } = 0m;
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
            // 1. Prepare sample data.
            List<Order> allOrders = new()
            {
                new Order { Id = 1, CustomerName = "Alice",   Status = "Pending",   Amount = 120.50m },
                new Order { Id = 2, CustomerName = "Bob",     Status = "Completed", Amount = 75.00m },
                new Order { Id = 3, CustomerName = "Charlie", Status = "Pending",   Amount = 200.00m },
                new Order { Id = 4, CustomerName = "Diana",   Status = "Shipped",   Amount = 50.25m }
            };

            // 2. Apply LINQ Where to keep only pending orders.
            List<Order> pendingOrders = allOrders
                .Where(o => o.Status == "Pending")
                .ToList();

            // 3. Populate the wrapper model.
            ReportModel model = new()
            {
                Orders = pendingOrders
            };

            // 4. Create the LINQ Reporting template programmatically.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders with status 'Pending':");
            // Use the correct property name 'Orders' (case‑sensitive) in the foreach tag.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("- Id: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Amount: $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // 5. Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // 6. Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // 7. Build the report using the filtered data.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 8. Save the final report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
