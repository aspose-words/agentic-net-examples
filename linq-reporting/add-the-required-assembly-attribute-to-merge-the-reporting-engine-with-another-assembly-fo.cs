using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Initialize properties to avoid nullable warnings.
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    // Collection that will be iterated in the template.
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");
        // Write a line for each person.
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new() { Name = "Alice", Age = 30 },
                new() { Name = "Bob", Age = 45 },
                new() { Name = "Charlie", Age = 28 }
            }
        };

        // Step 3: Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            // No special options needed for this simple example.
            Options = ReportBuildOptions.None
        };

        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Step 4: Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
