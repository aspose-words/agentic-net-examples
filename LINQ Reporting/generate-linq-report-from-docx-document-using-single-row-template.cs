using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace Example
{
    class Program
    {
        static void Main()
        {
            // Load the single‑row template (the template contains tags like <<[ds.Name]>>)
            Document template = new Document("Template.docx");

            // Example data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Id = 1, Name = "Alice", Age = 30 },
                new Person { Id = 2, Name = "Bob",   Age = 25 },
                new Person { Id = 3, Name = "Carol", Age = 28 }
            };

            // Use LINQ to pick a single row that will populate the template.
            Person? selectedPerson = people
                .Where(p => p.Id == 2)
                .SingleOrDefault();   // Returns null if not found – the engine can handle missing data.

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, selectedPerson, "ds");

            // Save the populated report.
            template.Save("Report.docx");
        }
    }

    // Supporting POCO class used as the data source.
    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
