using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple POCO class that will be used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Wrapper class that exposes a collection property.
    // The template can reference the collection via the property name.
    public class ReportData
    {
        public List<Person> Persons { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains LINQ Reporting tags.
            // Example template tags:
            //   <<foreach [Persons]>>
            //       <<[Name]>> (<<[Age]>> years old)
            //   <<endforeach>>
            Document template = new Document(@"C:\Templates\ReportTemplate.docm");

            // Create a list of Person objects using LINQ.
            List<Person> people = Enumerable.Range(1, 5)
                .Select(i => new Person
                {
                    Name = $"Person {i}",
                    Age = 20 + i
                })
                .ToList();

            // Wrap the list in a container object so the template can reference it.
            ReportData data = new ReportData { Persons = people };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data);

            // Save the populated document. The format is inferred from the file extension.
            template.Save(@"C:\Output\GeneratedReport.docx");
        }
    }
}
