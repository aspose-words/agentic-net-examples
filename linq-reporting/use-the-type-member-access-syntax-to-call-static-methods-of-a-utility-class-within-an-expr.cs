using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Sample data model.
            var model = new ReportModel
            {
                Name = "World"
            };

            // Create a template document with a LINQ Reporting tag that calls a static method.
            const string templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Use type member access syntax (::) with the fully‑qualified type name.
            builder.Writeln($"Hello <<[LinqReportingExample.Utility::UpperCase(Name)]>>!");
            doc.Save(templatePath);

            // Load the template.
            var template = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };
            bool success = engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);

            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }

    // Simple data model.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }

    // Utility class with a static method used in the template.
    public static class Utility
    {
        public static string UpperCase(string input) =>
            input?.ToUpperInvariant() ?? string.Empty;
    }
}
