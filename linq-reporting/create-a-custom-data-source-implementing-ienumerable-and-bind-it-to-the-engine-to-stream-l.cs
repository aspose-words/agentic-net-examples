using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCustomDataSource
{
    // Simple data model representing a person.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Custom data source that streams a large number of Person objects lazily.
    public class LargePersonDataSource : IEnumerable<Person>
    {
        private readonly int _count;

        public LargePersonDataSource(int count) => _count = count;

        public IEnumerator<Person> GetEnumerator()
        {
            for (int i = 0; i < _count; i++)
            {
                yield return new Person
                {
                    Name = $"Person {i + 1}",
                    Age = 20 + (i % 30)
                };
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    // Wrapper class used as the root object for the LINQ Reporting engine.
    public class ReportModel
    {
        public IEnumerable<Person> Persons { get; set; }

        public ReportModel()
        {
            // Simulate a large dataset of 10,000 persons.
            Persons = new LargePersonDataSource(10_000);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Person Report");
            builder.Writeln("==============");
            builder.Writeln();
            // LINQ Reporting foreach tag to iterate over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for report generation.
            Document report = new Document(templatePath);

            // Prepare the root data model.
            ReportModel model = new ReportModel();

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
