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
                new() { Name = "Alice", Age = 25 },
                new() { Name = "Bob", Age = 17 },
                new() { Name = "Charlie", Age = 30 },
                new() { Name = "Diana", Age = 15 }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("List of adults (Age > 18):");
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > 18)]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var templateDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        templateDoc.Save(outputPath);
    }
}
