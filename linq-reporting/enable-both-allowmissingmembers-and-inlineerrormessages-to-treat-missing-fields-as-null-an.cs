using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a Customer property.
    public class ReportModel
    {
        public CustomerInfo Customer { get; set; } = new CustomerInfo();
    }

    public class CustomerInfo
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Existing field – will be replaced with actual data.
            builder.Writeln("Customer Name: <<[model.Customer.Name]>>");

            // Missing field – will be treated as null and an inline error message will be inserted.
            builder.Writeln("Missing Field: <<[model.MissingObject.Id]>>");

            // Missing collection – the foreach loop will be ignored (treated as empty) and no error will be thrown.
            builder.Writeln("<<foreach [item in model.MissingCollection]>>Item: <<[item]>> <</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for reporting.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Customer = new CustomerInfo { Name = "John Doe" }
                // Note: No MissingObject or MissingCollection members are defined.
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine with the required options.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // Treat missing members as null and embed syntax error messages.
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages,
                MissingMemberMessage = "Missing"
            };

            // Build the report. The root object name is "model" to match the tags in the template.
            bool success = engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Output the result of the build operation.
            Console.WriteLine($"Report generation successful: {success}");
            Console.WriteLine($"Template: {templatePath}");
            Console.WriteLine($"Report: {reportPath}");
        }
    }
}
