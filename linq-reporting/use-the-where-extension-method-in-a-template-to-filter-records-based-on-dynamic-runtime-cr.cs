using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
        public int MinAge { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                MinAge = 30,
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 25 },
                    new Person { Name = "Bob", Age = 35 },
                    new Person { Name = "Charlie", Age = 30 },
                    new Person { Name = "Diana", Age = 40 }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("<<foreach [p in Persons.Where(p => p.Age >= MinAge)]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // Load the template
            var doc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
