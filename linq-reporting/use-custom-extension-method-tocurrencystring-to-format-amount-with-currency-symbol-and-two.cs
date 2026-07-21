using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCurrencyExample
{
    // Extension method used in the template.
    public static class Extensions
    {
        // Formats a decimal value as a currency string with a dollar sign and two decimal places.
        public static string ToCurrencyString(this decimal amount) => $"${amount:F2}";
    }

    // Simple data model for the report.
    public class ReportModel
    {
        // Sample amount to be formatted.
        public decimal Amount { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The tag calls the extension method ToCurrencyString on the Amount property.
            builder.Writeln("Amount: <<[model.Amount.ToCurrencyString()]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document (required by the lifecycle rules).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel { Amount = 1234.56m };

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Allow the engine to resolve extension methods.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Register the static class that contains the extension method.
            engine.KnownTypes.Add(typeof(Extensions));

            // Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
