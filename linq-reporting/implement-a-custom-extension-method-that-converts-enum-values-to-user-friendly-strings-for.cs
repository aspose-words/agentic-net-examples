using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingEnumExtension
{
    // Enum representing order status.
    public enum OrderStatus
    {
        Pending,
        Shipped,
        Delivered,
        Cancelled
    }

    // Extension/helper class that provides a user‑friendly string for the enum.
    public static class EnumExtensions
    {
        public static string ToFriendlyString(OrderStatus status) => status switch
        {
            OrderStatus.Pending => "Pending Approval",
            OrderStatus.Shipped => "Shipped to Customer",
            OrderStatus.Delivered => "Delivered Successfully",
            OrderStatus.Cancelled => "Order Cancelled",
            _ => status.ToString()
        };
    }

    // Data model for an order.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = "";
        public OrderStatus Status { get; set; } = OrderStatus.Pending;
        public List<Item> Items { get; set; } = new();

        // Computed property used by the LINQ Reporting template.
        public string StatusFriendly => EnumExtensions.ToFriendlyString(Status);
    }

    // Data model for an order item.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; } = 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                Id = 1001,
                CustomerName = "Acme Corp",
                Status = OrderStatus.Shipped,
                Items = new List<Item>
                {
                    new Item { Name = "Widget", Quantity = 5 },
                    new Item { Name = "Gadget", Quantity = 2 }
                }
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();

            builder.Writeln("Order Report");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Status: <<[order.StatusFriendly]>>");
            builder.Writeln();
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>> (Qty: <<[item.Quantity]>>)");
            builder.Writeln("<</foreach>>");

            builder.Document.Save(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
