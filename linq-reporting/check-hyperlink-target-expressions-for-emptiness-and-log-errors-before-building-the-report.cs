using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HyperlinkValidationExample
{
    // Data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Hyperlink target (URL or bookmark name). Initialized to empty string to avoid nullable warnings.
        public string Url { get; set; } = string.Empty;

        // Text displayed for the hyperlink. Initialized to a default value.
        public string Text { get; set; } = "Link";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting link tag that uses the model's Url and Text.
            // Syntax must be exactly as required: <<link [model.Url] [model.Text]>>
            builder.Writeln("<<link [model.Url] [model.Text]>>");

            // Save the template to disk (required before building the report).
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                // Intentionally leave Url empty to trigger validation.
                Url = string.Empty,
                Text = "Visit Site"
            };

            // -----------------------------------------------------------------
            // 4. Validate hyperlink target expressions before building the report.
            // -----------------------------------------------------------------
            bool canBuild = true;
            if (string.IsNullOrWhiteSpace(model.Url))
            {
                Console.WriteLine("Error: Hyperlink target (Url) is empty or whitespace.");
                canBuild = false;
            }

            // -----------------------------------------------------------------
            // 5. Build the report using Aspose.Words LINQ Reporting engine (only if valid).
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages; // Example option usage.

            bool success = false;
            if (canBuild)
            {
                success = engine.BuildReport(reportDoc, model, "model");
            }
            else
            {
                Console.WriteLine("Skipping report generation due to invalid hyperlink target.");
            }

            // -----------------------------------------------------------------
            // 6. Save the generated (or empty) report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }
}
