using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Prepare a collection of objects to be sorted in the template.
        var people = new List<Person>
        {
            new Person { Name = "Alice",   Age = 30 },
            new Person { Name = "Bob",     Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };

        // Create a new document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Sorted People (by Age):");

        // Insert a Reporting Engine template that uses the built‑in OrderBy extension method.
        // The expression Model.OrderBy(p => p.Age) sorts the collection before iteration.
        // Each iteration writes the person's name and age on a new line.
        builder.Writeln("{{foreach p in Model.OrderBy(person => person.Age)}}{{p.Name}} - {{p.Age}}\\n{{end}}");

        // Build the report by supplying the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new { Model = people });

        // Save the resulting document.
        doc.Save("SortedPeople.docx");
    }

    // Simple POCO class used as the data model.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
