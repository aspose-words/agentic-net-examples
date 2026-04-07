using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model representing an order.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = "";
        public double Amount { get; set; } = 0.0;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a collection of orders to be used as the data source.
            List<Order> orders = new()
            {
                new Order { Id = 1, CustomerName = "Alice", Amount = 123.45 },
                new Order { Id = 2, CustomerName = "Bob",   Amount = 678.90 },
                new Order { Id = 3, CustomerName = "Carol", Amount = 250.00 }
            };

            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Orders Report");
            builder.Writeln("<<foreach [order in orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Amount: $<<[order.Amount]>>");
            builder.Writeln("<</foreach>>");

            // Build the report using the overload that specifies the collection name.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, orders, "orders");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
