using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts the input string to title case using the current culture.
    public static string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
    }
}

// Simple data model used by the report.
public class Person
{
    public string Name { get; set; } = string.Empty;
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
            Persons = new()
            {
                new Person { Name = "john doe" },
                new Person { Name = "jane smith" },
                new Person { Name = "alice johnson" }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags: iterate over model.Persons and apply the helper method.
        builder.Writeln("<<foreach [p in model.Persons]>>");
        builder.Writeln("<<[StringHelpers.ToTitleCase(p.Name)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(StringHelpers));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
