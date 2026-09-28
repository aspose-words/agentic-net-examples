using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingEmptyParagraphRemoval
{
    // Root data model for the report.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Service> Services { get; set; } = new();
    }

    public class Service
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new Order
                    {
                        CustomerName = "Alice",
                        Services = new List<Service>
                        {
                            new Service { Name = "Consulting" },
                            new Service { Name = "Support" }
                        }
                    },
                    new Order
                    {
                        CustomerName = "Bob",
                        Services = new List<Service>() // No services – will generate empty paragraphs.
                    }
                }
            };

            // Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin outer foreach over Orders.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            // Intentionally add an empty paragraph that should be removed.
            builder.Writeln("");
            // Begin inner foreach over Services.
            builder.Writeln("<<foreach [svc in order.Services]>>");
            builder.Writeln("- Service: <<[svc.Name]>>");
            builder.Writeln("<</foreach>>");
            // End outer foreach.
            builder.Writeln("<</foreach>>");

            // Configure the reporting engine to remove empty paragraphs.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report.
            bool success = engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            doc.Save("Output.docx");
        }
    }
}
