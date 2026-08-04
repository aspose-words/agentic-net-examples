using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HyperlinkLinqReporting
{
    // Simple data model for the report.
    public class ReportModel
    {
        public string Url { get; set; } = "https://www.example.com";
        public string DisplayText { get; set; } = "Visit Example";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that creates a hyperlink.
            // The tag syntax is: <<link [uriExpression] [displayTextExpression]>>
            builder.Writeln("<<link [model.Url] [model.DisplayText]>>");

            // Save the template to disk (required by the workflow).
            const string templatePath = "HyperlinkTemplate.docx";
            template.Save(templatePath);

            // 2. Load the template back (simulating a separate load step).
            Document loadedTemplate = new Document(templatePath);

            // 3. Prepare the data source.
            ReportModel model = new ReportModel();

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this simple example.
            engine.BuildReport(loadedTemplate, model, "model");

            // 5. Save the generated report.
            const string outputPath = "HyperlinkReport.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
