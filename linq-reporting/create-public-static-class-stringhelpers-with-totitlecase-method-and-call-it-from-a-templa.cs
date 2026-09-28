using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts a string to title case (first letter of each word capitalized).
    public static string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
    }
}

// Data model classes
public class Person
{
    public string Name { get; set; } = string.Empty;

    // Exposes the title‑cased version of the name for the template.
    public string TitleCaseName => StringHelpers.ToTitleCase(Name);
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple foreach loop that uses the TitleCaseName property
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.TitleCaseName]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare sample data
        ReportModel model = new()
        {
            Persons = new List<Person>
            {
                new() { Name = "john doe" },
                new() { Name = "jane SMITH" },
                new() { Name = "alice o'connor" }
            }
        };

        // Build the report using LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);

        // Indicate completion (no interactive prompts)
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
