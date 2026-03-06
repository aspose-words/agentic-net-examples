using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains reporting tags (e.g. <<[persons.FullName]>>)
        Document template = new Document("Template.docx");

        // Sample data source: a list of Person objects
        List<Person> people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson", Age = 40 }
        };

        // Use LINQ to shape the data as needed by the template.
        // Here we project each Person to an anonymous type with properties matching the template tags.
        var reportData = people
            .Select(p => new
            {
                FullName = p.Name,
                Years = p.Age
            })
            .ToList();

        // Build the report. The third argument is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, reportData, "persons");

        // Save the populated document.
        template.Save("Report.docx");
    }

    // Simple POCO class used as the original data source.
    class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
