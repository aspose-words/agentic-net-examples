using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Disable reflection optimization for small data sets.
            ReportingEngine.UseReflectionOptimization = false;

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templatePath = "template.docx";
            var reportPath = "report.docx";

            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Person Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template so that it can be loaded before building the report.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            var person = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // -----------------------------------------------------------------
            // 3. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, person, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // Inform that the process completed (no interactive input required).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
