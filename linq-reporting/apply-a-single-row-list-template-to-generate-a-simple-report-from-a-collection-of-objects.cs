using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data – a single‑row list.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 }
            }
        };

        // Create a blank document and insert LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Simple Person Report");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the root data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}

// Root data model referenced in the template as <<[model]>>.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple item class used in the list.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
