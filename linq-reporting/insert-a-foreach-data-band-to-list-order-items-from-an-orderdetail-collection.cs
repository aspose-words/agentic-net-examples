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
        public string CustomerName { get; set; } = "";
        public List<OrderDetail> OrderDetails { get; set; } = new();
    }

    public class OrderDetail
    {
        public int Index { get; set; }
        public string ProductName { get; set; } = "";
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var order = new Order
            {
                CustomerName = "John Doe",
                OrderDetails = new List<OrderDetail>
                {
                    new OrderDetail { Index = 1, ProductName = "Apple", Quantity = 5, Price = 0.5m },
                    new OrderDetail { Index = 2, ProductName = "Banana", Quantity = 3, Price = 0.3m },
                    new OrderDetail { Index = 3, ProductName = "Cherry", Quantity = 10, Price = 0.2m }
                }
            };

            // Create template document programmatically
            const string templatePath = "OrderTemplate.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Order Report");
            builder.Writeln($"Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach data band for order details
            builder.Writeln("<<foreach [detail in OrderDetails]>>");

            // Create table header
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Qty");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Table row for each detail
            builder.InsertCell();
            builder.Writeln("<<[detail.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[detail.ProductName]>>");
            builder.InsertCell();
            builder.Writeln("<<[detail.Quantity]>>");
            builder.InsertCell();
            builder.Writeln("<<[detail.Price]>>");
            builder.EndRow();

            builder.EndTable();

            // End foreach block
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // Load the template for report generation
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report
            const string outputPath = "OrderReport.docx";
            doc.Save(outputPath);
        }
    }
}
