using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace LinqReportingExternalTypesExample
{
    // Model class that aggregates data from different namespaces.
    public class ReportModel
    {
        public List<Company.Models.Customer> Customers { get; set; } = new();
        public List<Company.Extra.Product> Products { get; set; } = new();
    }

    // First external type in its own namespace.
    namespace Company.Models
    {
        public class Customer
        {
            public string Name { get; set; } = "Unnamed";
            public int Age { get; set; }
        }
    }

    // Second external type in a different namespace.
    namespace Company.Extra
    {
        public class Product
        {
            public string Title { get; set; } = "Untitled";
            public decimal Price { get; set; }
        }
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for Aspose.Words (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Customers = new List<Company.Models.Customer>
                {
                    new() { Name = "Alice", Age = 30 },
                    new() { Name = "Bob", Age = 45 }
                },
                Products = new List<Company.Extra.Product>
                {
                    new() { Title = "Laptop", Price = 999.99m },
                    new() { Title = "Smartphone", Price = 499.50m }
                }
            };

            // Create a template document programmatically.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customers:");
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("- <<[c.Name]>> (<<[c.Age]>> years)");
            builder.Writeln("<</foreach>>");

            builder.Writeln();
            builder.Writeln("Products:");
            builder.Writeln("<<foreach [p in Products]>>");
            builder.Writeln("- <<[p.Title]>>: $<<[p.Price]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Configure the reporting engine.
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
