using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts; // Needed for ChartType enum

namespace LinkTagExample
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // URL that the link will point to.
        public string Url { get; set; } = "https://www.example.com";

        // Text displayed for the link.
        public string LinkText { get; set; } = "Visit Example";

        // Additional property to demonstrate that the model can hold more data.
        public string Title { get; set; } = "Link Tag Demo";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Add a heading (plain text, not a link).
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Insert a paragraph that contains a link tag.
            // The link tag will be resolved by the ReportingEngine.
            builder.Writeln("Here is a link: <<link [model.Url] [model.LinkText]>>");
            builder.Writeln();

            // Insert a chart to demonstrate that link tags are NOT placed inside chart elements.
            // The chart is added after the paragraph, so the link tag remains outside the chart.
            builder.InsertChart(ChartType.Column, 400, 300);
            builder.Writeln(); // Ensure a paragraph after the chart.

            // Save the template to disk.
            const string templatePath = "LinkTemplate.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create and configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name must match the tag prefix used in the template.
            engine.BuildReport(report, model, "model");

            // Save the generated report.
            const string outputPath = "LinkReport.docx";
            report.Save(outputPath);
        }
    }
}
