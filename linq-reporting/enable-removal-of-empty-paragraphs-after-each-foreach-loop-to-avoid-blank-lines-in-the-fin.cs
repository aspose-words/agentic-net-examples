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
        // Prepare sample data – one entry has an empty name to generate an empty line after the loop.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new() { Name = "Alice", Age = 30 },
                new() { Name = string.Empty, Age = 0 },   // Will produce an empty paragraph.
                new() { Name = "Bob", Age = 25 }
            }
        };

        // Create a temporary folder for the template and the generated report.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        string templatePath = Path.Combine(workDir, "Template.docx");
        string reportPath   = Path.Combine(workDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Person Report");
        builder.Writeln();

        // Insert a foreach block that iterates over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        // Each iteration writes a line with the person's name and age.
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        var engine = new ReportingEngine();
        // Enable removal of empty paragraphs that may appear after the foreach loop.
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save(reportPath);
    }
}
