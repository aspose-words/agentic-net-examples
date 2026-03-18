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

class Program
{
    static void Main()
    {
        // Load a template document that contains a placeholder like {{StringHelpers.ToTitleCase(Name)}}
        Document template = new Document("template.docx");

        // Simple anonymous object that provides the data used in the template.
        var data = new { Name = "john doe" };

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
