using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingIsolationExample
{
    // Sample data models
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<Service> Services { get; set; } = new();
    }

    public class Service
    {
        public string Name { get; set; } = "";
    }

    class Program
    {
        static void Main()
        {
            // Ensure output directory exists
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Path for the template document
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");

            // Create the LINQ Reporting template programmatically
            CreateTemplate(templatePath);

            // Simulate two independent user requests
            var orderUserA = new Order
            {
                CustomerName = "Alice Johnson",
                Services = new List<Service>
                {
                    new Service { Name = "Consultation" },
                    new Service { Name = "Implementation" }
                }
            };

            var orderUserB = new Order
            {
                CustomerName = "Bob Smith",
                Services = new List<Service>
                {
                    new Service { Name = "Support" },
                    new Service { Name = "Maintenance" },
                    new Service { Name = "Upgrade" }
                }
            };

            // Process each request with its own ReportingEngine instance
            ProcessRequest(orderUserA, templatePath, Path.Combine(outputDir, "Report_A.docx"));
            ProcessRequest(orderUserB, templatePath, Path.Combine(outputDir, "Report_B.docx"));
        }

        // Creates a simple Word template containing LINQ Reporting tags
        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Services:");
            builder.Writeln("<<foreach [svc in order.Services]>>");
            builder.Writeln("- <<[svc.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }

        // Generates a report for a single request using an isolated ReportingEngine
        private static void ProcessRequest(Order order, string templatePath, string outputPath)
        {
            // Load the template document
            var doc = new Document(templatePath);

            // Each request gets its own ReportingEngine instance
            var engine = new ReportingEngine();

            // Build the report using the order object as the root with name "order"
            engine.BuildReport(doc, order, "order");

            // Save the generated report
            doc.Save(outputPath);
        }
    }
}
