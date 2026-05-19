using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a simple DOCX template with a LINQ Reporting tag if it does not exist.
            const string templatePath = "Template.docx";
            if (!System.IO.File.Exists(templatePath))
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);
                builder.Writeln("Hello, <<[person.Name]>>!");
                doc.Save(templatePath);
            }

            // Load the template from the project directory.
            var template = new Document(templatePath);

            // Prepare the data model.
            var person = new Person();

            // Instantiate the reporting engine and build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(template, person, "person");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
