using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCurrencyExample
{
    // Sample data model.
    public class Order
    {
        public decimal Amount { get; set; } = 0m; // Initialize to avoid nullable warnings.
    }

    // Extension method used in the LINQ Reporting template.
    public static class Extensions
    {
        // Formats a decimal value as a currency string, e.g. $1234.56
        public static string ToCurrencyString(this decimal amount)
        {
            return string.Format("${0:N2}", amount);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document with a LINQ tag.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The tag calls the custom extension method ToCurrencyString().
            builder.Writeln("Amount: <<[order.Amount.ToCurrencyString()]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare data.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Sample data.
            Order order = new Order { Amount = 1234.56m };

            // -------------------------------------------------
            // 3. Build the report using ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Allow the engine to resolve extension methods.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(Extensions));

            // Build the report. The root name "order" must match the tag.
            engine.BuildReport(loadedTemplate, order, "order");

            // -------------------------------------------------
            // 4. Save the generated report.
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
