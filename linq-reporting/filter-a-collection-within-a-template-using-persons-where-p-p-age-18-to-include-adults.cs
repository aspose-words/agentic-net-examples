using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    // Collection that will be filtered inside the template.
    public List<Person> persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Prepare sample data.
        var model = new ReportModel
        {
            persons = new List<Person>
            {
                new Person { Name = "Alice",   Age = 17 },
                new Person { Name = "Bob",     Age = 22 },
                new Person { Name = "Charlie", Age = 19 },
                new Person { Name = "Diana",   Age = 15 }
            }
        };

        // 2. Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Adults (Age > 18):");

        // LINQ Reporting tag that filters the collection using Where().
        builder.Writeln("<<foreach [p in persons.Where(p => p.Age > 18)]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // 3. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // No special options are required for this simple example.
        engine.BuildReport(doc, model, "model");

        // 4. Save the generated report.
        doc.Save("AdultsReport.docx");
    }
}
