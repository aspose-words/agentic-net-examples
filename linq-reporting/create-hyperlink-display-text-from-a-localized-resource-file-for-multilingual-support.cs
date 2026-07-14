using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        public string Url { get; set; } = "";
        public string DisplayText { get; set; } = "";
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create simple localized text files (default and French).
        // -----------------------------------------------------------------
        string defaultTxt = Path.Combine(outputDir, "Strings.en.txt");
        string frenchTxt = Path.Combine(outputDir, "Strings.fr.txt");

        // Default (English) resources.
        File.WriteAllLines(defaultTxt, new[] { "LinkText=Click here" });

        // French resources.
        File.WriteAllLines(frenchTxt, new[] { "LinkText=Cliquez ici" });

        // -----------------------------------------------------------------
        // 2. Determine the UI culture and load the appropriate resource.
        // -----------------------------------------------------------------
        // For demonstration, switch to French culture; comment out to use default.
        CultureInfo.CurrentUICulture = new CultureInfo("fr-FR");

        string cultureSuffix = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "fr"
            ? ".fr"
            : ".en";

        string resourcePath = Path.Combine(outputDir, $"Strings{cultureSuffix}.txt");
        string localizedLinkText = LoadStringFromFile(resourcePath, "LinkText");

        // -----------------------------------------------------------------
        // 3. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a simple paragraph containing a link tag.
        // The tag uses the model's Url and DisplayText properties.
        builder.Writeln("<<link [model.Url] [model.DisplayText]>>");

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Url = "https://example.com",
            DisplayText = localizedLinkText
        };

        // -----------------------------------------------------------------
        // 5. Generate the report using ReportingEngine.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.docx");
        reportDoc.Save(resultPath);
    }

    // Helper method to read a key/value pair from a simple text file.
    private static string LoadStringFromFile(string filePath, string key)
    {
        if (!File.Exists(filePath))
            return string.Empty;

        foreach (var line in File.ReadAllLines(filePath))
        {
            var parts = line.Split('=', 2);
            if (parts.Length == 2 && parts[0].Trim() == key)
                return parts[1].Trim();
        }

        return string.Empty;
    }
}
