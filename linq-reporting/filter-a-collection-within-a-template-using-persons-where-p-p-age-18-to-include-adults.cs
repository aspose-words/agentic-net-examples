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
                new() { Name = "Alice", Age = 20 },
                new() { Name = "Bob", Age = 17 },
                new() { Name = "Charlie", Age = 25 }
            }
        };

        // Create the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("List of adults (Age > 18):");
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > 18)]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, can be omitted).
        const string templatePath = "template.docx";
        doc.Save(templatePath);

        // Load the template for report generation.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        templateDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success
            ? $"Report generated successfully: {outputPath}"
            : "Report generation failed.");
    }
}
