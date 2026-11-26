using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model.
    public class Order
    {
        // Initialize to avoid nullable warnings.
        public string CustomerName { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
    }

    // Custom utility class whose static methods will be used as extension methods in the template.
    public static class MyUtilities
    {
        // Extension method for DateTime to format the date.
        public static string FormatDate(this DateTime date)
        {
            return date.ToString("yyyy-MM-dd");
        }

        // Additional utility method (example) that could be used in templates.
        public static string ToUpperCase(this string text)
        {
            return text?.ToUpperInvariant() ?? string.Empty;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a simple template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags that reference the data model and the extension methods.
            builder.Writeln("Customer: <<[order.CustomerName.ToUpperCase()]>>");
            builder.Writeln("Order Date: <<[order.OrderDate.FormatDate()]>>");

            // Prepare sample data.
            Order order = new Order
            {
                CustomerName = "John Doe",
                OrderDate = DateTime.Now
            };

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Allow the engine to treat missing members (including extension methods) as valid.
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // Register the custom utility class so its static methods can be used as extension methods.
            engine.KnownTypes.Add(typeof(MyUtilities));

            // Build the report using the template, the data source, and the root name "order".
            engine.BuildReport(template, order, "order");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
