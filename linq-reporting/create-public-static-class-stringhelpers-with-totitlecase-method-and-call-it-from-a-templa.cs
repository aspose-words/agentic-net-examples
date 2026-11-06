using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class StringHelpers
{
    // Converts a string to title case (first letter of each word capitalized).
    public string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
        return textInfo.ToTitleCase(input.ToLower());
    }
}

// Data model classes
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
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document programmatically
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a heading
        builder.Writeln("Person List:");

        // Begin foreach loop over Persons collection
        builder.Writeln("<<foreach [p in Persons]>>");
        // Use StringHelpers.ToTitleCase to format the name
        builder.Writeln("Name: <<[StringHelpers.ToTitleCase(p.Name)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // 3. Prepare sample data
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "john doe" },
                new Person { Name = "jane smith" },
                new Person { Name = "alice johnson" }
            }
        };

        // 4. Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();

        // Provide both the model and the helper class instance as data sources
        engine.BuildReport(
            reportDoc,
            new object[] { model, new StringHelpers() },
            new[] { "model", "StringHelpers" });

        // 5. Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
