using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // In‑memory store for localized strings: culture code → (key → value)
    private static readonly Dictionary<string, Dictionary<string, string>> s_Resources = new();

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create simple resource entries for English and French.
        CreateResourceEntry("en", "LinkText", "Click here");
        CreateResourceEntry("fr", "LinkText", "Cliquez ici");

        // Build the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        BuildTemplate(templatePath);

        // Generate reports for two cultures.
        GenerateReport(outputDir, templatePath, "en-US", "Report_en.docx");
        GenerateReport(outputDir, templatePath, "fr-FR", "Report_fr.docx");
    }

    // Stores a single string entry for a given two‑letter ISO language.
    private static void CreateResourceEntry(string isoLanguage, string key, string value)
    {
        if (!s_Resources.TryGetValue(isoLanguage, out var dict))
        {
            dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            s_Resources[isoLanguage] = dict;
        }

        dict[key] = value;
    }

    // Builds a Word document containing a LINQ Reporting link tag.
    private static void BuildTemplate(string templatePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The template expects a root object named "model" with Url and LinkText properties.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        doc.Save(templatePath);
    }

    // Generates a report for a specific culture.
    private static void GenerateReport(string outputDir, string templatePath, string cultureName, string outputFileName)
    {
        // Load the template.
        Document template = new Document(templatePath);

        // Prepare the model data.
        ReportModel model = new ReportModel
        {
            Url = "https://example.com",
            LinkText = GetLocalizedString(cultureName, "LinkText")
        };

        // Build the report using LINQ Reporting.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        string outputPath = Path.Combine(outputDir, outputFileName);
        template.Save(outputPath);
    }

    // Retrieves a localized string from the in‑memory resource store.
    private static string GetLocalizedString(string cultureName, string key)
    {
        string iso = new CultureInfo(cultureName).TwoLetterISOLanguageName;
        if (s_Resources.TryGetValue(iso, out var dict) && dict.TryGetValue(key, out var value))
            return value;

        return string.Empty;
    }

    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        public string Url { get; set; } = string.Empty;
        public string LinkText { get; set; } = string.Empty;
    }
}
