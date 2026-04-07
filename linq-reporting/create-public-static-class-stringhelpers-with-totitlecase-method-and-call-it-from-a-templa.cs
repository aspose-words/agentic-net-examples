using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts a string to title case (first letter of each word capitalized).
    public static string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // Use the current culture for title casing.
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
    }
}

// Simple data model used as the root object for the report.
public class Person
{
    // Initialize to avoid nullable warnings.
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

        // Write a line that shows the original name.
        builder.Writeln("Original name: <<[person.Name]>>");
        // Write a line that uses the static helper to convert the name to title case.
        builder.Writeln("Title case: <<[StringHelpers.ToTitleCase(person.Name)]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Sample data.
        Person person = new Person { Name = "john doe" };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the static helper class so its methods can be used in the template.
        engine.KnownTypes.Add(typeof(StringHelpers));

        // Build the report. The root object name is "person".
        engine.BuildReport(loadedTemplate, person, "person");

        // Save the generated report.
        loadedTemplate.Save(reportPath);

        // Inform the user where the report was saved.
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
