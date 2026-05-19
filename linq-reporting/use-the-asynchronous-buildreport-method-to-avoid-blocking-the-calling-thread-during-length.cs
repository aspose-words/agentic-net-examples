using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingAsync
{
    // Simple data model used by the template.
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        // Async entry point.
        public static async Task Main(string[] args)
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Order Report");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- <<[item]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario).
            // -----------------------------------------------------------------
            Document doc = new(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            Order sampleOrder = new()
            {
                CustomerName = "John Doe",
                Items = new List<string> { "Apples", "Bananas", "Cherries" }
            };

            // -----------------------------------------------------------------
            // 4. Build the report asynchronously.
            // -----------------------------------------------------------------
            ReportingEngine engine = new();

            // Run the synchronous BuildReport method on a background thread to avoid blocking.
            bool success = await Task.Run(() => engine.BuildReport(doc, sampleOrder, "order"));

            // Optional: check the success flag when InlineErrorMessages option is used.
            if (!success)
            {
                Console.WriteLine("Report generation completed with errors.");
            }

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
