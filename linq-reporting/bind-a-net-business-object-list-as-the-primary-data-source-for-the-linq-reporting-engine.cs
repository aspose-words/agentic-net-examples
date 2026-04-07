using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Business object representing an order.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = string.Empty;
        public decimal Amount { get; set; } = 0m;
    }

    // Wrapper class that exposes the collection to the LINQ reporting engine.
    public class ReportData
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            List<Order> orders = new()
            {
                new Order { Id = 1, CustomerName = "Alice", Amount = 123.45m },
                new Order { Id = 2, CustomerName = "Bob",   Amount = 678.90m },
                new Order { Id = 3, CustomerName = "Carol", Amount = 250.00m }
            };

            // 2. Create the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Order Report");
            builder.Writeln("==============");
            // LINQ Reporting foreach tag.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Amount:   <<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required by the lifecycle rule).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 3. Load the template back (ensures the engine works on a loaded document).
            Document loadedTemplate = new Document(templatePath);

            // 4. Bind the data source.
            ReportData data = new()
            {
                Orders = orders
            };

            // 5. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed.
            engine.BuildReport(loadedTemplate, data);

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}
