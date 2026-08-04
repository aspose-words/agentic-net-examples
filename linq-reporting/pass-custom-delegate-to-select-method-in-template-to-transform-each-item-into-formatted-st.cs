using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Persons]>>");
        // Call the custom delegate (method) defined on the root model.
        builder.Writeln("<<[model.Format(p)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var loadedDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the final report.
        loadedDoc.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();

    // Custom delegate method used in the template to format a Person.
    public string Format(Person p) => $"{p.Name} (Age: {p.Age})";
}

// Simple data entity.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
