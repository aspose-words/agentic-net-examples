using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Extension-like static method that converts a string to upper case.
    public static string ToUpper(string value) => value?.ToUpper() ?? string.Empty;
}

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

        // -------------------------------------------------
        // Create the template document programmatically.
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that uses the static utility method.
        builder.Writeln("Original Name: <<[person.Name]>>");
        // Call the static method via its type name – method calls with '.' are not supported.
        builder.Writeln("Uppercase Name: <<[MyUtilities.ToUpper(person.Name)]>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -------------------------------------------------
        // Load the template document for reporting.
        // -------------------------------------------------
        var template = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the reporting engine and register the utility class.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        // Register the type that contains the static method.
        engine.KnownTypes.Add(typeof(MyUtilities));
        // Allow the engine to resolve static members (optional but safe).
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Sample data source.
        var person = new Person { Name = "John Doe" };

        // Build the report. The root object name must match the tag prefix.
        engine.BuildReport(template, person, "person");

        // Save the generated report.
        template.Save(reportPath);
    }
}
