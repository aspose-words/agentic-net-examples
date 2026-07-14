using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the template.
    public class ReportModel
    {
        public string Url { get; set; } = "https://example.com";
        public string LinkText { get; set; } = "Example Site";
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create the template document with a LINQ Reporting link tag.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Visit our site:");
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Step 2: Load the template for reporting.
            var doc = new Document(templatePath);

            // Step 3: Prepare the data source.
            var model = new ReportModel
            {
                Url = "https://www.aspose.com",
                LinkText = "Aspose.Words"
            };

            // Step 4: Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Step 5: Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
