using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Enum representing order status.
    public enum OrderStatus
    {
        Pending,
        Shipped,
        Delivered,
        Cancelled
    }

    // Static helper class that provides a user‑friendly string for each enum value.
    public static class EnumExtensions
    {
        // Note: This method is static so it can be called from a LINQ Reporting template.
        public static string ToFriendlyString(OrderStatus status)
        {
            return status switch
            {
                OrderStatus.Pending => "Pending Approval",
                OrderStatus.Shipped => "Shipped to Destination",
                OrderStatus.Delivered => "Delivered Successfully",
                OrderStatus.Cancelled => "Order Cancelled",
                _ => status.ToString()
            };
        }
    }

    // Simple data model for an order.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public OrderStatus Status { get; set; }
        public List<Item> Items { get; set; } = new();
    }

    // Simple data model for an item.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                CustomerName = "John Doe",
                Status = OrderStatus.Shipped,
                Items = new List<Item>
                {
                    new Item { Name = "Laptop", Quantity = 1 },
                    new Item { Name = "Mouse", Quantity = 2 }
                }
            };

            // Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            // Call the static helper method directly from the template.
            builder.Writeln("Status: <<[EnumExtensions.ToFriendlyString(order.Status)]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item.Name]>> x <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            // Register the static class that contains the helper method.
            engine.KnownTypes.Add(typeof(EnumExtensions));

            // The root object name must match the name used in the template tags.
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save("Report.docx", SaveFormat.Docx);
        }
    }
}
