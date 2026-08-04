using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Extension method to format a decimal as currency (e.g., $1234.56)
    public static class CurrencyExtensions
    {
        public static string ToCurrencyString(this decimal amount)
        {
            // Ensure two decimal places and prepend the dollar sign.
            return $"${amount:F2}";
        }
    }

    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public decimal Amount { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The tag uses the custom extension method ToCurrencyString().
            builder.Writeln("Amount: <<[model.Amount.ToCurrencyString()]>>");

            // Save the template to disk before building the report.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document (required before BuildReport).
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Amount = 1234.56m
            };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting Engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Allow the engine to resolve the extension method.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(CurrencyExtensions));

            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
