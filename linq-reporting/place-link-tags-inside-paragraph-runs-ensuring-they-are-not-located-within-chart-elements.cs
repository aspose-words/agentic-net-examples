using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // URL that the link will point to.
        public string Url { get; set; } = string.Empty;

        // Text that will be displayed for the link.
        public string LinkText { get; set; } = string.Empty;
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

            // Add a title.
            builder.Writeln("LINQ Reporting – Link Tag Example");
            builder.Writeln();

            // Insert a chart – the link must NOT be placed inside this chart.
            // The chart is added solely to demonstrate that the link is outside of it.
            builder.InsertChart(ChartType.Column, 400, 300);
            builder.Writeln(); // Ensure the cursor moves out of the chart.

            // Insert the LINQ Reporting link tag.
            // The tag uses the expressions [model.Url] and [model.LinkText] from the data source.
            builder.Writeln("Visit the website:");
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            var document = new Document(templatePath);

            // Prepare sample data.
            var model = new ReportModel
            {
                Url = "https://example.com",
                LinkText = "Example Site"
            };

            // Configure the reporting engine.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };

            // Build the report using the model as the root data source named "model".
            bool success = engine.BuildReport(document, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            document.Save(outputPath);
        }
    }
}
