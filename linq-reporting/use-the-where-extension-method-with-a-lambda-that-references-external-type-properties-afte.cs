using System;
using System.Collections.Generic;
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
                new() { Name = "Alice", Age = 28 },
                new() { Name = "Bob", Age = 35 },
                new() { Name = "Charlie", Age = 42 },
                new() { Name = "Diana", Age = 23 }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Filtered Persons (Age > 30):");
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > 30)]>>");
        builder.Writeln("- <<[p.Name]>> (<<[p.Age]>> years)");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // No explicit RegisterType call is required for this example.
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("Report.docx");
    }
}
