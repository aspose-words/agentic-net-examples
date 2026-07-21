using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new Model
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 }
            }
        };

        // Create a blank document and a builder to insert template tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Use the default LINQ Reporting delimiters (<< and >>).
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("ReportWithDefaultDelimiters.docx");
    }
}
