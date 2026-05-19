using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Create a template document containing a LINQ Reporting tag that references DateTime.UtcNow.
            const string templateFile = "Template.docx";
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Current UTC time: <<[DateTime.UtcNow]>>");
            template.Save(templateFile);

            // Load the template for report generation.
            Document report = new Document(templateFile);

            // Set up the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // Register the DateTime type so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(DateTime));

            // Build the report. No root data object is required because the template only uses a static expression.
            engine.BuildReport(report, new object());

            // Save the generated report.
            const string outputFile = "Report.docx";
            report.Save(outputFile);
        }
    }
}
