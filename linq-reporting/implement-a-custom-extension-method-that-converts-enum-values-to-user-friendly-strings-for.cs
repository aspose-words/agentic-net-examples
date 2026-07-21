using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample enum representing order status.
    public enum OrderStatus
    {
        Pending,
        Shipped,
        Delivered
    }

    // Extension methods for enums.
    public static class EnumExtensions
    {
        // Converts an OrderStatus value to a user‑friendly string.
        public static string ToFriendlyString(this OrderStatus status)
        {
            return status switch
            {
                OrderStatus.Pending => "Pending Approval",
                OrderStatus.Shipped => "Shipped Out",
                OrderStatus.Delivered => "Delivered to Customer",
                _ => status.ToString()
            };
        }
    }

    // Data model used by the LINQ Reporting engine.
    public class Order
    {
        public int Id { get; set; } = 0;
        public OrderStatus Status { get; set; } = OrderStatus.Pending;
        public string CustomerName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                Id = 123,
                CustomerName = "John Doe",
                Status = OrderStatus.Shipped
            };

            // Create a temporary folder for the template and result.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            string templatePath = Path.Combine(workDir, "Template.docx");
            string resultPath = Path.Combine(workDir, "Result.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Order Report");
            builder.Writeln("--------------------");
            builder.Writeln($"Order ID: <<[order.Id]>>");
            builder.Writeln($"Customer: <<[order.CustomerName]>>");
            // Use a static call to the extension method (supported by the engine).
            builder.Writeln($"Status: <<[EnumExtensions.ToFriendlyString(order.Status)]>>");

            // Save the template.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(EnumExtensions));

            // Build the report using the root object name "order".
            engine.BuildReport(loadedDoc, order, "order");

            // Save the generated report.
            loadedDoc.Save(resultPath);
        }
    }
}
