using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSignatureExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Report title – non‑nullable to avoid CS8618.
        public string Title { get; set; } = string.Empty;

        // Signatory name – nullable; when null the signature block is omitted.
        public string? Signatory { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a title placeholder.
            builder.Writeln("Report Title: <<[model.Title]>>");

            // Conditional block – the content inside will be included only if
            // model.Signatory is not null.
            builder.Writeln("<<if [model.Signatory != null]>>");
            builder.Writeln("Signature: ______________________");
            builder.Writeln("Signed by: <<[model.Signatory]>>");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            var model = new ReportModel
            {
                Title = "Monthly Sales Report",
                Signatory = "John Doe" // Change to null to omit the signature block.
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // The root object name must match the name used in the template tags.
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
