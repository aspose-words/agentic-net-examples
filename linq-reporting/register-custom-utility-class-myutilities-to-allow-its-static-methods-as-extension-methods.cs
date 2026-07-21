using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Custom utility class with an extension method.
    public static class MyUtilities
    {
        // Extension method for DateTime to format as short date string.
        public static string ToShortDate(this DateTime date)
        {
            return date.ToString("yyyy-MM-dd");
        }
    }

    // Simple data model used as the root object for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public DateTime OrderDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                CustomerName = "John Doe",
                OrderDate = new DateTime(2023, 5, 21)
            };

            // Create a template document programmatically.
            string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template document.
            var doc = new Document(templatePath);

            // Configure the reporting engine.
            var engine = new ReportingEngine();

            // Register the custom utility class so its static/extension methods can be used in the template.
            engine.KnownTypes.Add(typeof(MyUtilities));

            // Allow the engine to ignore missing members (optional, but safe).
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // Build the report using the root object name "order".
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            string reportPath = "Report.docx";
            doc.Save(reportPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }

        // Helper method to create the template with LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert tags that reference the data model and use the extension method.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order Date: <<[order.OrderDate.ToShortDate()]>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
