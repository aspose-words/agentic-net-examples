using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple resource dictionary for two languages.
        var resources = new Dictionary<string, Dictionary<string, string>>
        {
            ["en"] = new Dictionary<string, string>
            {
                ["LinkText"] = "Visit Aspose",
                ["Url"] = "https://www.aspose.com"
            },
            ["fr"] = new Dictionary<string, string>
            {
                ["LinkText"] = "Visitez Aspose",
                ["Url"] = "https://www.aspose.com/fr"
            }
        };

        // Choose a language (in a real scenario this could come from user settings).
        string language = "fr";

        // Build the data model that the LINQ Reporting engine will use.
        var model = new ReportModel
        {
            Url = resources[language]["Url"],
            LinkText = resources[language]["LinkText"]
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a link tag where the first expression is the URL and the second
        // expression is the display text taken from the data model.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template (simulating a separate load step).
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final document.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model exposed to the template. All properties are non‑nullable to avoid warnings.
public class ReportModel
{
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}
