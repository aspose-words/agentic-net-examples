using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Sample data model.
    public class Order
    {
        // Initialize to avoid nullable warnings.
        public decimal Total { get; set; } = 0m;
    }

    // Custom external type whose static members can be used in the template.
    public static class MyHelper
    {
        // Formats a decimal value as currency.
        public static string FormatCurrency(decimal value)
        {
            return string.Format("{0:C}", value);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that calls the custom static method.
            // The root data source will be referenced as "order".
            builder.Writeln("Order total: <<[MyHelper.FormatCurrency(order.Total)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data.
            Order order = new Order { Total = 1234.56m };

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register the custom external type so the template can access its static members.
            engine.KnownTypes.Add(typeof(MyHelper));

            // Build the report. The root object name must match the name used in the template tags.
            engine.BuildReport(reportDoc, order, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
