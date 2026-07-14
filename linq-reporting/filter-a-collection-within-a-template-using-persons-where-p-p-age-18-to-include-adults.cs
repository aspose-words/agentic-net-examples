using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model representing a person.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Adults (Age > 18):");
            // The foreach tag filters the collection directly in the template.
            builder.Writeln("<<foreach [p in model.Persons.Where(p => p.Age > 18)]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice",   Age = 20 },
                    new Person { Name = "Bob",     Age = 17 },
                    new Person { Name = "Charlie", Age = 25 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // The root name "model" must match the name used in the template tags.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
