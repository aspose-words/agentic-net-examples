using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used in the report.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Hello, <<[person.Name]>>!");

            // Save the template to disk (optional, demonstrates load/save workflow).
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // Load the template back (simulating a real-world scenario where the template is read from a file).
            Document doc = new Document(templatePath);

            // Specify prohibited .NET types before any report generation.
            ReportingEngine.SetRestrictedTypes(typeof(System.Environment), typeof(System.IO.File));

            // Prepare the data source.
            Person person = new Person();

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);
        }
    }
}
