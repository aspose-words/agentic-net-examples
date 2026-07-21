using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a collection of persons.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document containing LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Template iterates over the Persons collection and writes each person's data.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template document for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 40 }
                }
            };

            // 4. Configure the ReportingEngine to inline error messages.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // 5. Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);

            // Output the result of the build operation.
            Console.WriteLine($"Report build successful: {success}");
        }
    }
}
