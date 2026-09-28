using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    // Note: No Address property – it will be missing in the data model.
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Address: <<[p.Address]>>"); // Address does not exist in Person.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new Model
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 }
            }
        };

        // Configure the reporting engine to allow missing members.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
