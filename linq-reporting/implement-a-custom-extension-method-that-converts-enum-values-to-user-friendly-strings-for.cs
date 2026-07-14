using System;
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

    // Extension method that converts enum values to user‑friendly strings.
    public static class OrderStatusExtensions
    {
        public static string ToFriendlyString(this OrderStatus status)
        {
            return status switch
            {
                OrderStatus.Pending   => "Pending Approval",
                OrderStatus.Shipped   => "Shipped to Customer",
                OrderStatus.Delivered => "Delivered Successfully",
                OrderStatus.Cancelled => "Order Cancelled",
                _                     => status.ToString()
            };
        }
    }

    // Data model used by the LINQ Reporting engine.
    public class Order
    {
        public int Id { get; set; }
        public OrderStatus Status { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document with a LINQ Reporting tag.
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Order Status: <<[order.Status.ToFriendlyString()]>>");
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template (simulating a separate load step).
            var loadedTemplate = new Document(templatePath);

            // 3. Prepare the data source.
            var order = new Order
            {
                Id = 12345,
                Status = OrderStatus.Shipped
            };

            // 4. Configure the ReportingEngine.
            var engine = new ReportingEngine();

            // Enable extension method resolution.
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(OrderStatusExtensions));

            // 5. Build the report.
            engine.BuildReport(loadedTemplate, order, "order");

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
