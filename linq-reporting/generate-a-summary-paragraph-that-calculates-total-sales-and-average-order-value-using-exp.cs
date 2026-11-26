using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSummary
{
    // Root data model
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    // Order item
    public class Order
    {
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create template document programmatically
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Sales Summary");
            builder.Writeln("Total Sales: <<[model.Orders.Sum(o => o.Amount)]>>");
            builder.Writeln("Average Order Value: <<[model.Orders.Average(o => o.Amount)]>>");

            // Save template to disk
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Prepare sample data
            var model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new() { Amount = 120.50m },
                    new() { Amount = 340.00m },
                    new() { Amount = 215.75m }
                }
            };

            // Build the report using LINQ Reporting Engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            // Indicate completion (no interactive input)
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
