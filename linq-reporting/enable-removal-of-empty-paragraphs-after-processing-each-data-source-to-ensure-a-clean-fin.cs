using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data where some descriptions are empty.
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Description = "Software engineer" },
                new Person { Name = "Bob", Description = "" },               // Empty description
                new Person { Name = "Charlie", Description = "Project manager" },
                new Person { Name = "Diana", Description = null }           // Null description
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employee Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Description: <<[p.Description]>>"); // May become empty
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to remove empty paragraphs after processing.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        doc.Save("EmployeeReport.docx");
    }
}

// Wrapper class for the data source.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data entity.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public string? Description { get; set; }
}
