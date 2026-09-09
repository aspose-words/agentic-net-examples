using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = string.Empty;
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
        // Prepare sample data (unsorted).
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Charlie", Age = 30 },
                new Person { Name = "Alice",   Age = 25 },
                new Person { Name = "Bob",     Age = 28 }
            }
        };

        // Create the LINQ Reporting template.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People sorted by name (inline sorting):");
        builder.Writeln("<<foreach [p in Persons.OrderBy(p => p.Name)]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
