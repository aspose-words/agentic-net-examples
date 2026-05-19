using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class StringHelpers
{
    // Converts a string to title case (first letter of each word capitalized).
    public static string ToTitleCase(string value)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        // Use the current culture's TextInfo for title casing.
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
    }
}

// Simple data model used as the root object for the report.
public class Model
{
    public string Name { get; set; } = "john doe";
}

public class Program
{
    public static void Main()
    {
        // Create a blank Word document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that calls the static ToTitleCase method.
        // The root object is referenced as "model", so we use model.Name as the argument.
        builder.Writeln("<<[StringHelpers.ToTitleCase(model.Name)]>>");

        // Prepare the data source.
        Model model = new Model();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the helper class so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(StringHelpers));

        // Build the report using the template, data source, and root name.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
