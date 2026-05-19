using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalRemoval
{
    // Data model used by the LINQ Reporting template.
    public class ReportModel
    {
        public bool ShowA { get; set; } = false;
        public bool ShowB { get; set; } = false;
        public string Title { get; set; } = "Conditional Block Demo";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title paragraph.
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln(); // Empty line for visual separation.

            // Paragraph that contains two conditional blocks.
            // If neither condition is true, the paragraph becomes empty and will be removed.
            builder.Writeln("<<if [model.ShowA]>>Option A is shown.<</if>>");
            builder.Writeln("<<if [model.ShowB]>>Option B is shown.<</if>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                // Both booleans are false, so no conditional content should appear.
                ShowA = false,
                ShowB = false,
                Title = "Report with Conditional Removal"
            };

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after processing the conditional tags.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // The root object name in the template is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
