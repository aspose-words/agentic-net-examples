using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
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
            // Register code page provider (required for some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Step 1: Create the template document programmatically
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Simple Report");
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Step 2: Load the template (simulating a separate load step)
            var loadedTemplate = new Document(templatePath);

            // Step 3: Prepare sample data
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Step 4: Build the report using the LINQ Reporting engine
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required
            engine.BuildReport(loadedTemplate, model, "model");

            // Step 5: Save the generated report
            const string reportPath = "report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
