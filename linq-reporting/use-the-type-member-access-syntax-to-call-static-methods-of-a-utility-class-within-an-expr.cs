using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Utility class with a static method that will be called from a LINQ Reporting expression tag.
    public static class MyUtility
    {
        // Formats a DateTime value as a short date string.
        public static string FormatDate(DateTime date) => date.ToString("yyyy-MM-dd");
    }

    // Simple data model that will be used as the root object for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public DateTime OrderDate { get; set; } = DateTime.Today;
        public decimal Amount { get; set; } = 123.45m;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words for some encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            // Call static method MyUtility.FormatDate via expression tag.
            builder.Writeln("Order Date: <<[MyUtility.FormatDate(order.OrderDate)]>>");
            builder.Writeln("Amount: <<[order.Amount]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back from disk (required before building the report).
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var order = new Order
            {
                CustomerName = "Alice Smith",
                OrderDate = new DateTime(2023, 12, 15),
                Amount = 987.65m
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();

            // Register the utility type so its static members can be accessed in expressions.
            engine.KnownTypes.Add(typeof(MyUtility));

            // The root object name used in the template tags is "order".
            engine.BuildReport(loadedTemplate, order, "order");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
