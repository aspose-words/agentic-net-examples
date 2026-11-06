using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLambdaExample
{
    // Data model classes
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public decimal TotalAmount { get; set; }
    }

    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
        public decimal Threshold { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Threshold = 150m,
                Orders = new List<Order>
                {
                    new Order { CustomerName = "Alice", TotalAmount = 120m },
                    new Order { CustomerName = "Bob",   TotalAmount = 200m },
                    new Order { CustomerName = "Carol", TotalAmount = 300m }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("Orders with total greater than <<[model.Threshold]>>:");
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.TotalAmount > model.Threshold)]>>");
            builder.Writeln("- <<[order.CustomerName]>>: $<<[order.TotalAmount]>>");
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
