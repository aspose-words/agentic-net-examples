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
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob",   Age = 35 },
                new Person { Name = "Carol", Age = 42 },
                new Person { Name = "Dave",  Age = 31 },
                new Person { Name = "Eve",   Age = 24 }
            }
        };

        // Create a blank document and insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a single LINQ expression that combines Where, Select, and OrderBy.
        // The expression filters persons aged 30 or more, selects their names,
        // and orders the names alphabetically.
        builder.Writeln("<<foreach [name in Persons.Where(p => p.Age >= 30).Select(p => p.Name).OrderBy(n => n)]>>");
        builder.Writeln("<<[name]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
