using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // LINQ Reporting tags: iterate over the Persons collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data. Person objects do NOT have an Age property.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice" },
                new Person { Name = "Bob" }
            }
        };

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A";

        // Build the report. The root object name is "model".
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportOutput.docx");
    }
}

// Wrapper class that serves as the root data source.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Data class with only a Name property; Age is intentionally missing.
public class Person
{
    public string Name { get; set; } = string.Empty;
}
