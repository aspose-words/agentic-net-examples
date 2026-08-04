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

        // Use the current culture for proper casing.
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
    }
}

public class Model
{
    // Sample property that will be transformed by the helper.
    public string Name { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string resultPath = "Result.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that calls the static helper method.
        builder.Writeln("Hello <<[StringHelpers.ToTitleCase(Name)]>>!");

        // Save the template so it can be loaded for reporting.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare data source.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);
        var model = new Model { Name = "john doe" };

        // -------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();

        // Register the helper class so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(StringHelpers));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(resultPath);
    }
}
