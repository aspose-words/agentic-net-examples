using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains a numbered‑list placeholder.
        // Example placeholder in the template:
        // <<foreach [people]>><<[Name]>><</foreach>>
        Document template = new Document("Template.docx"); // create/load rule

        // Sample data source.
        var people = new List<Person>
        {
            new Person { Name = "Alice",   Age = 30 },
            new Person { Name = "Bob",     Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // LINQ query – order the collection by Age descending.
        var orderedPeople = people.OrderBy(p => -p.Age).ToList();

        // Populate the template with the LINQ result.
        ReportingEngine engine = new ReportingEngine();
        // The data source name ("people") must match the name used in the template.
        engine.BuildReport(template, orderedPeople, "people"); // build report rule

        // Save the generated report.
        template.Save("Report.docx"); // save rule
    }

    // Simple POCO used as the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
