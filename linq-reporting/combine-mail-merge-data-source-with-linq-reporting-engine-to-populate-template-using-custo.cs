using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingMailMergeExample
{
    // Public data model classes
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Address { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a template document programmatically
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Report");
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("Name: <<[c.Name]>>");
            builder.Writeln("Email: <<[c.Email]>>");
            builder.Writeln("Address: <<[c.Address]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Prepare sample data
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new() { Name = "John Doe", Email = "john@example.com", Address = "123 Main St" },
                    new() { Name = "Jane Smith", Email = "jane@example.com", Address = "456 Oak Ave" }
                }
            };

            // Build the report using LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            var outputPath = "output.docx";
            reportDoc.Save(outputPath);
        }
    }
}
