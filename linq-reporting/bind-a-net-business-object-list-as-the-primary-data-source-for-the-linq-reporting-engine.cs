using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Business object representing a simple order.
    public class Order
    {
        public int Id { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public decimal Amount { get; set; }
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
            var orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice", Amount = 123.45m },
                new Order { Id = 2, CustomerName = "Bob",   Amount = 678.90m },
                new Order { Id = 3, CustomerName = "Carol", Amount = 250.00m }
            };

            var model = new ReportModel { Orders = orders };

            // 2. Create a template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Order Report");
            builder.Writeln("==============");
            builder.Writeln(); // empty line

            // Begin a foreach block that iterates over the Orders collection.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Amount:   <<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // 3. Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 4. Load the template (simulating a separate load step).
            var loadedTemplate = new Document(templatePath);

            // 5. Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options

            // The root object name in the template is "model", so we pass it accordingly.
            engine.BuildReport(loadedTemplate, model, "model");

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
