using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingExample
{
    // Data model classes
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer { Name = "Alice Johnson" },
                    new Customer { Name = "Bob Smith" },
                    new Customer { Name = "Charlie Davis" }
                }
            };

            // Create a template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer List:");
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln(" - <<[c.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model"
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
