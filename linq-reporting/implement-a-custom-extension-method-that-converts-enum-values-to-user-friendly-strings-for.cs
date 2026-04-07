using System;
using System.IO;
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
    public static class EnumExtensions
    {
        public static string ToFriendlyString(this OrderStatus status)
        {
            return status switch
            {
                OrderStatus.Pending => "Pending",
                OrderStatus.Shipped => "Shipped",
                OrderStatus.Delivered => "Delivered",
                OrderStatus.Cancelled => "Cancelled",
                _ => status.ToString()
            };
        }
    }

    // Data model used by the LINQ Reporting engine.
    public class Order
    {
        public int Id { get; set; } = 1001;
        public string CustomerName { get; set; } = "John Doe";
        public OrderStatus Status { get; set; } = OrderStatus.Pending;

        // Property that uses the extension method for template display.
        public string FriendlyStatus => Status.ToFriendlyString();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Status: <<[order.FriendlyStatus]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Sample data source.
            Order sampleOrder = new Order
            {
                Id = 12345,
                CustomerName = "Alice Smith",
                Status = OrderStatus.Shipped
            };

            // Use the ReportingEngine to populate the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, sampleOrder, "order");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // Inform the user where the report was saved.
            Console.WriteLine($"Report generated at: {Path.GetFullPath(reportPath)}");
        }
    }
}
