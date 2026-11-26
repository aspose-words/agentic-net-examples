using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize properties to avoid nullable warnings.
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Set the UI culture to Spanish to demonstrate localization.
        CultureInfo.CurrentUICulture = new CultureInfo("es");

        // Retrieve the localized display text for the hyperlink.
        string linkText = GetLocalizedString("LinkText");

        // Prepare the data model for the report.
        var model = new ReportModel
        {
            Url = "https://example.com",
            LinkText = linkText
        };

        // Build the LINQ Reporting template.
        const string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Generate the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Simple in‑memory localization dictionary (replaces runtime .resx handling).
    private static string GetLocalizedString(string key)
    {
        // Define resources for each culture (empty key = invariant/default).
        var resources = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            { "", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "LinkText", "Click here" } } },
            { "es", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "LinkText", "Haga clic aquí" } } }
        };

        // Use the two‑letter ISO language name (e.g., "es") for lookup.
        string culture = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;

        // Try to get the value for the specific culture; fall back to invariant.
        if (resources.TryGetValue(culture, out var dict) && dict.TryGetValue(key, out var value))
            return value;

        // Fallback to invariant resources.
        if (resources[""].TryGetValue(key, out var defaultValue))
            return defaultValue;

        return string.Empty;
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a hyperlink where the display text comes from the model.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        doc.Save(path);
    }
}
