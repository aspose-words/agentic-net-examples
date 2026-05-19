using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // External type with a static property used in the LINQ filter.
    public static class FilterHelper
    {
        // Minimum age for filtering persons.
        public static int MinAge { get; set; } = 30;
    }

    // Data model representing a person.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Root data source for the report.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 25 },
                    new Person { Name = "Bob",   Age = 35 },
                    new Person { Name = "Carol", Age = 45 }
                }
            };

            // 2. Create a template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Use Where extension method with a lambda that references the external static property.
            builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > FilterHelper.MinAge)]>>");
            builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 3. Load the template for report generation.
            var doc = new Document(templatePath);

            // 4. Configure the ReportingEngine.
            var engine = new ReportingEngine();
            // Register the external type so its static members can be used in expressions.
            engine.KnownTypes.Add(typeof(FilterHelper));

            // 5. Build the report. Use the overload without a root name to reference members directly.
            engine.BuildReport(doc, model);

            // 6. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
