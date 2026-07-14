using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // External static class whose property will be used inside the LINQ expression.
    public static class FilterSettings
    {
        // Minimum age to include in the report.
        public static int MinAge { get; set; } = 30;
    }

    // Data model class.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Person> persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("<<foreach [p in persons.Where(p => p.Age > FilterSettings.MinAge)]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // 2. Load the template.
            var doc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 25 },
                    new Person { Name = "Bob", Age = 35 },
                    new Person { Name = "Charlie", Age = 45 }
                }
            };

            // 4. Configure the reporting engine.
            var engine = new ReportingEngine();
            // Register the external type so its static members can be accessed in the template.
            engine.KnownTypes.Add(typeof(FilterSettings));

            // 5. Build the report.
            bool success = engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);

            // Optional: indicate success (no console input required).
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }
}
