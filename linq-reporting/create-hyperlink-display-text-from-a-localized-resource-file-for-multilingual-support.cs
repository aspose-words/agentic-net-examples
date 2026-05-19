using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Set the culture to French to demonstrate localized text.
        CultureInfo.CurrentCulture = new CultureInfo("fr-FR");

        // -----------------------------------------------------------------
        // 1. Create a simple JSON resource file with translations.
        // -----------------------------------------------------------------
        string resourcesPath = "resources.json";
        var translations = new Dictionary<string, string>
        {
            { "en", "Visit our site" },
            { "fr", "Visitez notre site" }
        };
        File.WriteAllText(resourcesPath, JsonConvert.SerializeObject(translations));

        // -----------------------------------------------------------------
        // 2. Load the localized display text based on the current culture.
        // -----------------------------------------------------------------
        var loadedTranslations = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(resourcesPath));
        string languageKey = CultureInfo.CurrentCulture.TwoLetterISOLanguageName;
        string displayText = loadedTranslations.ContainsKey(languageKey) ? loadedTranslations[languageKey] : loadedTranslations["en"];

        // -----------------------------------------------------------------
        // 3. Prepare the data model for the report.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Url = "https://www.example.com",
            DisplayText = displayText
        };

        // -----------------------------------------------------------------
        // 4. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        // Insert a link tag where the display text comes from the model.
        builder.Writeln("<<link [model.Url] [model.DisplayText]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 5. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    public string Url { get; set; } = "";
    public string DisplayText { get; set; } = "";
}
