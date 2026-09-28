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
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create the template document with inline sorting using LINQ extension methods.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Sorted Persons (by Name):");
        builder.Writeln("<<foreach [person in model.Persons.OrderBy(p => p.Name)]>>");
        builder.Writeln("<<[person.Name]>> - <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template to simulate a real file scenario.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);
        var loadedTemplate = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output file: {outputPath}");
    }
}
