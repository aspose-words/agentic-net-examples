using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Title will be displayed in the report.
        public string Title { get; set; } = string.Empty;

        // This property is intentionally left null to demonstrate removal of empty paragraphs.
        public string? EmptyTag { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Paragraph that will be kept (contains a non‑empty value).
            builder.Writeln("<<[model.Title]>>");

            // Paragraph that contains only a tag whose value is null/empty.
            // After the report is built this paragraph becomes empty and should be removed.
            builder.Writeln("<<[model.EmptyTag]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Configure the reporting engine to remove empty paragraphs.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Prepare the data source.
            var model = new ReportModel
            {
                Title = "Sample Report"
                // EmptyTag remains null.
            };

            // Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);
        }
    }
}
