using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Data model classes
    public class ReportData
    {
        public string CustomerName { get; set; } = "Acme Corp";
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public int Id { get; set; }
        public decimal Total { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document with LINQ Reporting tags
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Report");
            builder.Writeln("Customer: <<[data.CustomerName]>>");
            builder.Writeln("Orders:");
            builder.Writeln("<<foreach [order in data.Orders]>>");
            builder.Writeln("- Order ID: <<[order.Id]>>, Total: $<<[order.Total]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template for report generation
            var templateDoc = new Document(templatePath);

            // Prepare sample data
            var data = new ReportData
            {
                Orders = new List<Order>
                {
                    new Order { Id = 101, Total = 250.75m },
                    new Order { Id = 102, Total = 89.99m },
                    new Order { Id = 103, Total = 1234.00m }
                }
            };

            // Enable reflection optimization for better performance on large collections
            ReportingEngine.UseReflectionOptimization = true;

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, data, "data");

            // Save the generated report
            var outputPath = "Report.docx";
            templateDoc.Save(outputPath);
        }
    }
}
