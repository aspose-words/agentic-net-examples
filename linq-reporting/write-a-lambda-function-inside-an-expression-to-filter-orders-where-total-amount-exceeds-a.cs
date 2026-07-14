using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingLambdaFilter
{
    // Data model for an order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public decimal TotalAmount { get; set; }
    }

    // Wrapper model that holds the collection and the threshold.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
        public decimal Threshold { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Threshold = 100m,
                Orders = new List<Order>
                {
                    new Order { CustomerName = "Alice", TotalAmount = 75m },
                    new Order { CustomerName = "Bob",   TotalAmount = 150m },
                    new Order { CustomerName = "Carol", TotalAmount = 200m },
                    new Order { CustomerName = "Dave",  TotalAmount = 50m }
                }
            };

            // 2. Create the template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Orders with total amount greater than <<[model.Threshold]>>:");
            // Use a lambda expression inside the foreach tag to filter the collection.
            builder.Writeln("<<foreach [order in model.Orders.Where(o => o.TotalAmount > model.Threshold)]>>");
            builder.Writeln("- <<[order.CustomerName]>>: $<<[order.TotalAmount]>>");
            builder.Writeln("<</foreach>>");

            // Save the template before building the report.
            doc.Save(templatePath);

            // 3. Load the template and build the report.
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(loadedDoc, model, "model");

            // 4. Save the generated report.
            loadedDoc.Save("Report.docx");
        }
    }
}
