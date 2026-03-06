using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains reporting tags, e.g. <<[ds.Name]>> and <<[ds.Salary]>>
        Document doc = new Document("Template.docx");

        // Sample data source: a list of Person objects
        List<Person> people = new List<Person>
        {
            new Person { Name = "Alice",   Age = 28, Salary = 50000 },
            new Person { Name = "Bob",     Age = 35, Salary = 70000 },
            new Person { Name = "Charlie", Age = 42, Salary = 90000 }
        };

        // LINQ query: filter persons older than 30 and project only the needed fields
        var query = people
            .Where(p => p.Age > 30)
            .Select(p => new { p.Name, p.Salary })
            .ToList();

        // Build the report. The data source name "ds" must match the tags in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, query, "ds");

        // Save the populated document.
        doc.Save("Report.docx");
    }
}

// Simple POCO representing a person.
class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
    public double Salary { get; set; }
}
