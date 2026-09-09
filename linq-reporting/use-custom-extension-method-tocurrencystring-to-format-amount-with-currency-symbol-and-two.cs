using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Extension method to format a decimal as currency with a dollar sign and two decimal places.
    public static class Extensions
    {
        public static string ToCurrencyString(this decimal amount) => $"${amount:0.00}";
    }

    // Simple data model used as the root object for the report.
    public class Order
    {
        public decimal Amount { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order { Amount = 1234.567m };

            // Create a temporary folder for the template and output files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinqReportingDemo");
            Directory.CreateDirectory(workDir);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "Template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that calls the custom extension method.
            // The tag references the root object name "order".
            builder.Writeln("Amount: <<[order.Amount.ToCurrencyString()]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine
            {
                // Allow the engine to use extension methods defined in the project.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report using the "order" root name.
            engine.BuildReport(doc, order, "order");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(workDir, "Report.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
