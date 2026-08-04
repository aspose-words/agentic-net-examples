using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample enum representing order status.
    public enum OrderStatus
    {
        Pending,
        Shipped,
        Delivered,
        Cancelled
    }

    // Extension method (also usable as a static method) to convert enum values to user‑friendly strings.
    public static class OrderStatusExtensions
    {
        // This method can be called as an extension method or as a static method.
        public static string ToFriendlyString(this OrderStatus status)
        {
            return status switch
            {
                OrderStatus.Pending => "Pending Approval",
                OrderStatus.Shipped => "Shipped to Customer",
                OrderStatus.Delivered => "Delivered Successfully",
                OrderStatus.Cancelled => "Order Cancelled",
                _ => "Unknown"
            };
        }
    }

    // Data model for a single order.
    public class Order
    {
        public int Id { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public OrderStatus Status { get; set; }
    }

    // Wrapper model that contains a collection of orders.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            // Use the static method call syntax that the ReportingEngine can resolve.
            builder.Writeln("Status: <<[OrderStatusExtensions.ToFriendlyString(order.Status)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Prepare sample data.
            var orders = new List<Order>
            {
                new Order { Id = 1, CustomerName = "Alice", Status = OrderStatus.Pending },
                new Order { Id = 2, CustomerName = "Bob",   Status = OrderStatus.Shipped },
                new Order { Id = 3, CustomerName = "Carol", Status = OrderStatus.Delivered },
                new Order { Id = 4, CustomerName = "Dave",  Status = OrderStatus.Cancelled }
            };

            var model = new ReportModel { Orders = orders };

            // Step 3: Load the template and build the report.
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Register the type that contains the static method so the engine can invoke it.
            engine.KnownTypes.Add(typeof(OrderStatusExtensions));

            // Build the report using the wrapper model; the root name is "model".
            engine.BuildReport(doc, model, "model");

            // Step 4: Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
