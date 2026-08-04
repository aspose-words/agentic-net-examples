using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Drawing.Charts; // ChartType resides in this namespace

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL that the link will point to.
        public string Url { get; set; } = "https://example.com";

        // Text displayed for the link.
        public string Text { get; set; } = "Visit Example.com";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph that will contain the link tag.
            builder.Writeln("Please click the following link:");
            // Write the link tag inside the same paragraph run (no line break).
            builder.Write("<<link [model.Url] [model.Text]>>");

            // Insert a chart after the paragraph to demonstrate that the link is NOT inside a chart.
            // The chart itself does not contain any LINQ Reporting tags.
            builder.Writeln(); // Ensure the chart starts on a new paragraph.
            builder.InsertChart(ChartType.Column, 400, 300);

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure and execute the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            bool success = engine.BuildReport(reportDoc, model, "model");

            // (Optional) You can check the success flag if InlineErrorMessages were enabled.
            if (!success)
            {
                Console.WriteLine("Report generation encountered errors.");
            }

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath, SaveFormat.Docx);
        }
    }
}
