using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingEnumExtensionExample
{
    // Enum representing order status.
    public enum OrderStatus
    {
        Pending,
        InProgress,
        Completed
    }

    // Extension method to convert enum values to user‑friendly strings.
    public static class EnumExtensions
    {
        public static string ToFriendlyString(this OrderStatus status) =>
            status switch
            {
                OrderStatus.Pending => "Pending Approval",
                OrderStatus.InProgress => "In Progress",
                OrderStatus.Completed => "Completed Successfully",
                _ => status.ToString()
            };
    }

    // Model class for a single order.
    public class Order
    {
        public int Id { get; set; }
        public OrderStatus Status { get; set; }

        // Expose a user‑friendly status string for the template.
        public string StatusFriendly => Status.ToFriendlyString();

        // Parameterless constructor to satisfy nullable warnings.
        public Order() { }
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();

        // Parameterless constructor.
        public ReportModel() { }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new() { Id = 101, Status = OrderStatus.Pending },
                    new() { Id = 102, Status = OrderStatus.InProgress },
                    new() { Id = 103, Status = OrderStatus.Completed }
                }
            };

            // Paths for template and output documents.
            string templatePath = "template.docx";
            string outputPath = "report.docx";

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders Report");
            builder.Writeln(); // Empty line.

            // Begin foreach loop over Orders collection.
            builder.Writeln("<<foreach [order in model.Orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Status: <<[order.StatusFriendly]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and build the report.
            // -------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
