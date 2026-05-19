using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model representing a customer.
    public class Customer
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Add a title.
            builder.Writeln("Customer Names Report");
            builder.Writeln();

            // LINQ Reporting tags:
            //   <<foreach [c in Customers]>> ... <</foreach>>
            //   <<[c.Name]>> outputs the customer's name.
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("- <<[c.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required by the workflow).
            // -----------------------------------------------------------------
            var document = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer { Name = "Alice Johnson" },
                    new Customer { Name = "Bob Smith" },
                    new Customer { Name = "Charlie Davis" }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(document, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            document.Save(outputPath);
        }
    }
}
