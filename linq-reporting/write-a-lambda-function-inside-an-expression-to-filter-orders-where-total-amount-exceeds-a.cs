using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLambdaExample
{
    // Simple data model representing an order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public double TotalAmount { get; set; }
    }

    // Wrapper class that will be passed as the root data source to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new() { CustomerName = "Alice", TotalAmount = 750.0 },
                    new() { CustomerName = "Bob",   TotalAmount = 1250.0 },
                    new() { CustomerName = "Carol", TotalAmount = 300.0 },
                    new() { CustomerName = "Dave",  TotalAmount = 2100.0 }
                }
            };

            // 2. Create a Word template programmatically.
            const string templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Write a foreach tag that uses a lambda expression to filter orders
            // with a total amount greater than 1000.
            builder.Writeln("Orders with total amount > 1000:");
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.TotalAmount > 1000)]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>, Total: <<[order.TotalAmount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // 3. Load the template and build the report.
            var template = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(template, model, "model");

            // 4. Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
