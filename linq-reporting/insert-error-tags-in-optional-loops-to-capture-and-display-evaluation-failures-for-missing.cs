using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write a simple report with an optional loop.
        builder.Writeln("Report:");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // Age property does not exist in the data model; <<error>> will capture the evaluation failure.
        builder.Writeln("Age: <<[p.Age]>> <<error>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data. The Person class lacks an Age property.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice" },
                new Person { Name = "Bob" }
            }
        };

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages | ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missing"
        };

        // Build the report. The success flag is meaningful because InlineErrorMessages is enabled.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Output the result status (no interactive prompts required).
        Console.WriteLine($"Report generation succeeded: {success}");
    }
}

// Root data model referenced in the template as <<[model.Persons]>>.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Sample item class. Intentionally does NOT contain an Age property to trigger an error.
public class Person
{
    public string Name { get; set; } = "";
}
