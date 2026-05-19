using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for an order detail line.
    public class OrderDetail
    {
        public int Index { get; set; }
        public string ProductName { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
    }

    // Data model for an order containing a collection of order details.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<OrderDetail> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<OrderDetail>
                {
                    new OrderDetail { Index = 1, ProductName = "Apple", Quantity = 5, UnitPrice = 0.60m },
                    new OrderDetail { Index = 2, ProductName = "Banana", Quantity = 3, UnitPrice = 0.40m },
                    new OrderDetail { Index = 3, ProductName = "Cherry", Quantity = 10, UnitPrice = 0.15m }
                }
            };

            // Create a template document programmatically.
            var templatePath = "OrderTemplate.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Order for: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln(" - <<[item.Index]>>: <<[item.ProductName]>> x <<[item.Quantity]>> @ <<[item.UnitPrice]>>");
            builder.Writeln("<</foreach>>");

            // Save the template before building the report.
            templateDoc.Save(templatePath);

            // Load the template and build the report.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            var outputPath = "OrderReport.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
