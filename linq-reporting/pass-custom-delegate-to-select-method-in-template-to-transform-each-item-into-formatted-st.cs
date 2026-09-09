using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // The template iterates over a transformed collection using LINQ Select.
        // The lambda concatenates FirstName and LastName into a formatted string.
        builder.Writeln("<<foreach [fullName in Persons.Select(p => p.FirstName + \" \" + p.LastName)]>>");
        builder.Writeln("Name: <<[fullName]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { FirstName = "John", LastName = "Doe" },
                new Person { FirstName = "Jane", LastName = "Smith" },
                new Person { FirstName = "Bob",  LastName = "Johnson" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the final report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// Root data model referenced in the template as "model".
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data entity.
public class Person
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
}
