using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts a string to title case (e.g., "john doe" -> "John Doe").
    public static string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
        return textInfo.ToTitleCase(input.ToLower());
    }
}

public class Person
{
    // Sample property used in the template.
    public string Name { get; set; } = string.Empty;

    // Exposes the title‑cased version of Name using the helper.
    public string TitleCaseName => StringHelpers.ToTitleCase(Name);
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a line that uses the TitleCaseName property.
        builder.Writeln("Hello, <<[model.TitleCaseName]>>!");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        Person model = new Person { Name = "john doe" };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
