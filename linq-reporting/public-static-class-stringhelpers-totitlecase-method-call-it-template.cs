using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts the supplied text to title case (first letter of each word capitalized).
    public static string ToTitleCase(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
        // Ensure the whole string is lower‑cased before applying TitleCase to avoid culture‑specific quirks.
        return textInfo.ToTitleCase(input.ToLower());
    }
}

// Simple data class required by Aspose.Words.Reporting (must be a visible type).
public class Person
{
    public string Name { get; set; }
}

class Program
{
    static void Main()
    {
        // Create a template document in memory with a reporting placeholder.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("{{StringHelpers.ToTitleCase(Name)}}");

        // Provide the data source as a visible type instance.
        var data = new Person { Name = "john doe" };

        // Configure the ReportingEngine so it knows about the static helper class.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(StringHelpers));

        // Build the report – the engine will evaluate the expression in the template
        // and replace it with the title‑cased version of the Name field.
        engine.BuildReport(template, data);

        // Save the generated document.
        template.Save("output.docx");
    }
}
