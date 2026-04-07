using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank Word document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that accesses the static member Math.PI.
            // The tag <<[Math.PI]>> will be replaced with the value of Math.PI during report generation.
            builder.Writeln("The value of PI is <<[Math.PI]>>.");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back (simulating a real‑world scenario where the template is stored separately).
            Document doc = new Document(templatePath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register System.Math in the KnownTypes collection so that its static members can be used in tags.
            engine.KnownTypes.Add(typeof(System.Math));

            // Build the report. No data source is required because we only use static members.
            // The overload without a data source name is sufficient.
            engine.BuildReport(doc, new object());

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
