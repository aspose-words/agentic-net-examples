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
        // Ensure the output directory exists.
        Directory.CreateDirectory("output");

        // Simulate a .resx resource file with localized strings using a dictionary.
        var resources = new Dictionary<string, string>
        {
            ["LinkUrl"] = "https://example.com",
            ["LinkText_en"] = "Visit Example",
            ["LinkText_es"] = "Visitar Ejemplo"
        };

        // Set the UI culture to Spanish to demonstrate localization.
        CultureInfo.CurrentUICulture = new CultureInfo("es");

        // Determine the appropriate link text based on the current UI culture.
        string cultureKey = $"LinkText_{CultureInfo.CurrentUICulture.TwoLetterISOLanguageName}";
        string linkText = resources.TryGetValue(cultureKey, out var localizedText)
            ? localizedText
            : resources["LinkText_en"];

        // Prepare the data model for the report.
        var model = new ReportModel
        {
            Url = resources["LinkUrl"],
            LinkText = linkText
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = "output/template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Please click the link below:");
        // Insert the link tag using the model properties.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string outputPath = "output/ReportWithLocalizedLink.docx";
        doc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}
