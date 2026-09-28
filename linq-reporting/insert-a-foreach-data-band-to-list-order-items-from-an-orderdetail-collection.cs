using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    // Data model classes
    public class Order
    {
        public string OrderId { get; set; } = "";
        public string CustomerName { get; set; } = "";
        public List<OrderDetail> Items { get; set; } = new();
    }

    public class OrderDetail
    {
        public int Index { get; set; }
        public string ProductName { get; set; } = "";
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal Total => Quantity * UnitPrice;
    }

    public class Program
    {
        public static void Main()
        {
            // Sample data
            var order = new Order
            {
                OrderId = "ORD-001",
                CustomerName = "John Doe",
                Items = new List<OrderDetail>
                {
                    new OrderDetail { Index = 1, ProductName = "Apple", Quantity = 3, UnitPrice = 0.5m },
                    new OrderDetail { Index = 2, ProductName = "Banana", Quantity = 5, UnitPrice = 0.3m },
                    new OrderDetail { Index = 3, ProductName = "Orange", Quantity = 2, UnitPrice = 0.8m }
                }
            };

            // Create template document
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Order Report");
            builder.Writeln("Order ID: <<[order.OrderId]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach data band for order items
            builder.Writeln("<<foreach [item in order.Items]>>");

            // Table with header and data rows
            Table table = builder.StartTable();

            // Header row
            builder.InsertCell(); builder.Writeln("Index");
            builder.InsertCell(); builder.Writeln("Product");
            builder.InsertCell(); builder.Writeln("Qty");
            builder.InsertCell(); builder.Writeln("Unit Price");
            builder.InsertCell(); builder.Writeln("Total");
            builder.EndRow();

            // Data row (repeated for each item)
            builder.InsertCell(); builder.Writeln("<<[item.Index]>>");
            builder.InsertCell(); builder.Writeln("<<[item.ProductName]>>");
            builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
            builder.InsertCell(); builder.Writeln("<<[item.UnitPrice]>>");
            builder.InsertCell(); builder.Writeln("<<[item.Total]>>");
            builder.EndRow();

            builder.EndTable();

            // End foreach data band
            builder.Writeln("<</foreach>>");

            // Save the template
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Save the final report
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);
        }
    }
}
