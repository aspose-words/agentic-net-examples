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
                new() { Name = "Alice", Age = 30 },
                new() { Name = "Bob", Age = 25 },
                new() { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons sorted by Age (ascending):");
        builder.Writeln("<<foreach [p in Persons.OrderBy(p => p.Age)]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var loadDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        bool success = engine.BuildReport(loadDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        loadDoc.Save(reportPath);

        // Optional: indicate success (no console interaction required).
        // The program ends here.
    }
}
