using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Order
    {
        // Initialize to avoid nullable warnings.
        public string CustomerName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths in the current directory.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that references the root object named "order".
            builder.Writeln("Customer: <<[order.CustomerName]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            // Load the previously saved template.
            Document doc = new Document(templatePath);

            // Create sample data.
            Order order = new Order { CustomerName = "John Doe" };

            // Store the current setting so we can restore it later.
            bool originalOptimization = ReportingEngine.UseReflectionOptimization;

            try
            {
                // Disable reflection optimization for this specific report generation.
                ReportingEngine.UseReflectionOptimization = false;

                // Build the report using the LINQ Reporting engine.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(doc, order, "order");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = originalOptimization;
            }

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
