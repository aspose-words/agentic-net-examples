using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSignatureExample
{
    // Data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "";
        public string Content { get; set; } = "";
        // Nullable signatory – when null the signature block is omitted.
        public string? Signatory { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Title = "Quarterly Summary",
                Content = "All targets have been met for this quarter.",
                Signatory = "Jane Doe" // Set to null to omit the signature block.
            };

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Simple report fields.
            builder.Writeln("Report Title: <<[model.Title]>>");
            builder.Writeln("Report Content: <<[model.Content]>>");

            // Conditional block – the signature line appears only when Signatory is not null.
            builder.Writeln("<<if [model.Signatory != null]>>");
            builder.Writeln("Signed by: <<[model.Signatory]>>");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // BuildReport with the root object name "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);
        }
    }
}
