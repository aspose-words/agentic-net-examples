using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Non‑nullable properties are initialized to avoid warnings.
        public string Name { get; set; } = string.Empty;
        public string OptionalText { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains a placeholder which may be empty.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Paragraph with a required value.
            builder.Writeln("Customer: <<[model.Name]>>");

            // Paragraph that can become empty after rendering (OptionalText may be empty).
            builder.Writeln("Note: <<[model.OptionalText]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Populate the model – OptionalText is left empty to trigger removal.
            var model = new ReportModel
            {
                Name = "John Doe",
                OptionalText = string.Empty // This will result in an empty paragraph.
            };

            // Configure the ReportingEngine to remove empty paragraphs after processing.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the final document.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Optional: output a short confirmation.
            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
