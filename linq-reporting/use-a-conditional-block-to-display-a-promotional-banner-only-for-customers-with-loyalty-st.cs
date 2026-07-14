using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalBanner
{
    // Data model for a customer.
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public bool IsLoyal { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Customers collection.
            builder.Writeln("<<foreach [c in Customers]>>");
            // Output the customer's name.
            builder.Writeln("Customer: <<[c.Name]>>");
            // Conditional block: show promotional banner only for loyal customers.
            builder.Writeln("<<if [c.IsLoyal]>>");
            builder.Writeln("=== Promotional Banner: 20% OFF! ===");
            builder.Writeln("<</if>>");
            // End of foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare sample data.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            ReportModel model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer { Name = "Alice", IsLoyal = true },
                    new Customer { Name = "Bob",   IsLoyal = false },
                    new Customer { Name = "Carol", IsLoyal = true }
                }
            };

            // -------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated report.
            // -------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
