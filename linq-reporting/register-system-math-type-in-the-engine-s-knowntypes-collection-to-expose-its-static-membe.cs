using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Value of Math.PI: <<[Math.PI]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for reporting.
            Document report = new Document(templatePath);

            // Register System.Math so its static members can be used in tags.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(System.Math));

            // Build the report. No data source is required for this example.
            engine.BuildReport(report, new object());

            // Save the generated report.
            const string outputPath = "Report.docx";
            report.Save(outputPath);

            // Indicate completion (optional).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
