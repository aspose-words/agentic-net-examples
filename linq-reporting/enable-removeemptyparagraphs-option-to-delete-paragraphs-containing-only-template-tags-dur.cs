using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // This property will be null, causing the corresponding tag to render an empty string.
        public string? Optional { get; set; } = null;

        // Additional property to demonstrate that the report still contains content.
        public string Message { get; set; } = "Report generated successfully.";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Paragraph that contains only a LINQ Reporting tag.
            // Since ReportModel.Optional is null, this paragraph will become empty after processing.
            builder.Writeln("<<[model.Optional]>>");

            // Another paragraph with regular text to verify that the document still has content.
            builder.Writeln("<<[model.Message]>>");

            // Save the template to a local file (optional, shown for clarity).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to remove empty paragraphs after tag processing.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the template and the data model.
            // The root object name in the template is "model".
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
