using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts a string to title case (first letter of each word capitalized).
    public static string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return string.Empty;

        // Use the current culture for title casing.
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
    }
}

// Simple data model used by the report.
public class Person
{
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that calls the static helper method.
        // The tag will output the person's name in title case.
        builder.Writeln("Hello <<[StringHelpers.ToTitleCase(Name)]>>!");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data source.
        Person person = new Person { Name = "john doe" };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the helper class so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(StringHelpers));

        // Build the report. No root name is needed because we reference members directly.
        engine.BuildReport(reportDoc, person);

        // Save the generated report.
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
