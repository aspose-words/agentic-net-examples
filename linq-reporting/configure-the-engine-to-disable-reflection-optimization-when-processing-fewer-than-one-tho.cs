using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data entity.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model that matches the template root name.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data (fewer than 1000 records).
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 45 },
                    new Person { Name = "Charlie", Age = 28 }
                }
            };

            // Create a template document with LINQ Reporting tags.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            doc.Save(templatePath);

            // Load the template for reporting.
            var template = new Document(templatePath);

            // Disable reflection optimization when processing fewer than 1000 records.
            if (model.Persons.Count < 1000)
                ReportingEngine.UseReflectionOptimization = false;
            else
                ReportingEngine.UseReflectionOptimization = true; // default behavior

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
