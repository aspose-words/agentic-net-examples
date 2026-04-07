using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        const string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Prepare localized resources manually (no .resx file needed).
        // -----------------------------------------------------------------
        var resources = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            // Invariant (default) culture.
            { string.Empty, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "LinkText", "Visit Aspose" }
                }
            },
            // French culture.
            { "fr-FR", new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "LinkText", "Visitez Aspose" }
                }
            }
        };

        // -----------------------------------------------------------------
        // 2. Load the resource according to the current UI culture.
        // -----------------------------------------------------------------
        CultureInfo.CurrentUICulture = new CultureInfo("fr-FR");
        string displayText = ResolveResource(resources, "LinkText", CultureInfo.CurrentUICulture);

        // -----------------------------------------------------------------
        // 3. Prepare the data model that will be bound to the template.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Url = "https://www.aspose.com",
            DisplayText = displayText
        };

        // -----------------------------------------------------------------
        // 4. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Execute the reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);

        // Clean up temporary files (optional).
        File.Delete(templatePath);
    }

    // Retrieves a resource value for the given key respecting the supplied culture.
    private static string ResolveResource(
        Dictionary<string, Dictionary<string, string>> resources,
        string key,
        CultureInfo culture)
    {
        // Try the specific culture first.
        if (resources.TryGetValue(culture.Name, out var dict) && dict.TryGetValue(key, out var val))
            return val;

        // Fallback to parent culture.
        if (!culture.IsNeutralCulture && culture.Parent != CultureInfo.InvariantCulture)
            return ResolveResource(resources, key, culture.Parent);

        // Finally, fallback to invariant (default) value.
        if (resources.TryGetValue(string.Empty, out var invariantDict) && invariantDict.TryGetValue(key, out var invariantVal))
            return invariantVal;

        // If not found, return the key itself.
        return key;
    }

    // Creates a simple Word document containing a LINQ Reporting link tag.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // The template uses the model root name "model".
        builder.Writeln("<<link [model.Url] [model.DisplayText]>>");

        doc.Save(filePath);
    }

    // Data model used by the reporting engine.
    public class ReportModel
    {
        public string Url { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
    }
}
