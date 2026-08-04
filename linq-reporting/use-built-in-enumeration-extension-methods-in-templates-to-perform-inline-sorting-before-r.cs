using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSortingExample
{
    // Simple data model.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data (unsorted).
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 34 },
                    new Person { Name = "Bob",   Age = 28 },
                    new Person { Name = "Carol", Age = 45 },
                    new Person { Name = "Dave",  Age = 22 }
                }
            };

            // Create a template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("People sorted by Age (ascending):");
            // Inline sorting using LINQ OrderBy extension method.
            builder.Writeln("<<foreach [p in Persons.OrderBy(person => person.Age)]>>");
            builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
            builder.Writeln("<</foreach>>");

            // Save the template (optional, just to illustrate the lifecycle).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template (demonstrating load step).
            var loadedTemplate = new Document(templatePath);

            // Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the final report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
