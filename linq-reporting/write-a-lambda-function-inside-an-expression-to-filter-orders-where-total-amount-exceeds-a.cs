using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model representing an order.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = string.Empty;
        public decimal TotalAmount { get; set; } = 0m;
    }

    // Wrapper class that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
        public decimal Threshold { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Threshold = 150m,
                Orders = new List<Order>
                {
                    new Order { Id = 1, CustomerName = "Alice",   TotalAmount = 120m },
                    new Order { Id = 2, CustomerName = "Bob",     TotalAmount = 200m },
                    new Order { Id = 3, CustomerName = "Charlie", TotalAmount = 350m },
                    new Order { Id = 4, CustomerName = "Diana",   TotalAmount = 80m }
                }
            };

            // 2. Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // LINQ Reporting tag with a lambda expression that filters orders.
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.TotalAmount > model.Threshold)]>>");
            builder.Writeln("Order ID: <<[order.Id]>>, Customer: <<[order.CustomerName]>>, Total: <<[order.TotalAmount]>>");
            builder.Writeln("<</foreach>>");

            // 3. Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 4. Load the template back (required by the workflow).
            var loadedTemplate = new Document(templatePath);

            // 5. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
