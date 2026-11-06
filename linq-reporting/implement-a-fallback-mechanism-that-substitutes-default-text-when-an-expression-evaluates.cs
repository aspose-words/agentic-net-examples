using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string? Name { get; set; }
    public int Age { get; set; }

    public Person(string? name, int age)
    {
        Name = name;
        Age = age;
    }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data: one person with a null Name to trigger the fallback.
        var model = new ReportModel();
        model.Persons.Add(new Person(null, 30));
        model.Persons.Add(new Person("Alice", 25));

        // Create the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Person Report");
        builder.Writeln();

        // Begin a foreach loop over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        // Use null‑coalescing to substitute default text when Name is null.
        builder.Writeln("Name: <<[p.Name ?? \"(no name)\"]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
