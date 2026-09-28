using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingEnumExtension
{
    // Sample enum representing order status.
    public enum OrderStatus
    {
        Pending,
        Shipped,
        Delivered,
        Cancelled
    }

    // Extension method that converts enum values to user‑friendly strings.
    public static class OrderStatusExtensions
    {
        public static string ToFriendlyString(this OrderStatus status) => status switch
        {
            OrderStatus.Pending => "Pending Approval",
            OrderStatus.Shipped => "Shipped to Customer",
            OrderStatus.Delivered => "Delivered Successfully",
            OrderStatus.Cancelled => "Order Cancelled",
            _ => "Unknown"
        };
    }

    // Model class used as a data source for the report.
    public class ReportModel
    {
        public List<OrderItem> Items { get; set; } = new();
    }

    // Individual item containing an enum property.
    public class OrderItem
    {
        public string ProductName { get; set; } = "";
        public OrderStatus Status { get; set; }

        // Friendly string representation of the status for the template.
        public string StatusFriendly => Status.ToFriendlyString();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<OrderItem>
                {
                    new OrderItem { ProductName = "Laptop", Status = OrderStatus.Pending },
                    new OrderItem { ProductName = "Smartphone", Status = OrderStatus.Shipped },
                    new OrderItem { ProductName = "Headphones", Status = OrderStatus.Delivered },
                    new OrderItem { ProductName = "Monitor", Status = OrderStatus.Cancelled }
                }
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("Order Report");
            builder.Writeln("==============");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Product: <<[item.ProductName]>>");
            builder.Writeln("Status: <<[item.StatusFriendly]>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
