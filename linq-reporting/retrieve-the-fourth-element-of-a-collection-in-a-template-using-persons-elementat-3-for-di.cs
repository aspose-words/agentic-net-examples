using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 28 },
                new Person { Name = "Diana", Age = 32 },
                new Person { Name = "Ethan", Age = 27 }
            }
        };

        // Create a template document with a LINQ Reporting tag that retrieves the fourth element.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Fourth person in the list:");
        // Use the correct root name ("model") and property name ("Persons").
        builder.Writeln("<<[model.Persons.ElementAt(3).Name]>>");
        builder.Document.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the model as the root object named "model".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
