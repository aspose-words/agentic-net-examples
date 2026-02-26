using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains Reporting Engine tags (e.g. <<foreach [people]>> <<[Name]>> <<[Age]>> <</foreach>>)
        Document template = new Document("Template.docx");

        // Prepare a collection of data objects using LINQ.
        // In a real scenario the source could be a database query; here we use an in‑memory list.
        var people = new List<Person>
        {
            new Person { Name = "Alice",   Age = 30 },
            new Person { Name = "Bob",     Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // Example LINQ operation: order the collection by Age.
        var orderedPeople = people.OrderBy(p => p.Age).ToList();

        // Populate the template with the ordered collection.
        // The third argument ("people") is the name used inside the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, orderedPeople, "people");

        // Save the generated report as a DOCX file.
        template.Save("Result.docx");
    }

    // Simple POCO class that represents a row in the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
