using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRemoveEmptyParagraphs
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph with normal text and a tag.
            builder.Writeln("Hello, <<[model.Name]>>!");

            // Paragraph that contains only a tag – should be removed after build.
            builder.Writeln("<<[model.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template document for reporting.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Configure the reporting engine to remove empty paragraphs.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report using the model as the root object.
            ReportModel model = new ReportModel();
            engine.BuildReport(doc, model, "model");

            // Save the final report.
            doc.Save(reportPath);
        }
    }
}
