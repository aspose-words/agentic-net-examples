using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        Directory.CreateDirectory("output");

        // 1. Prepare localized link texts.
        var resources = new Dictionary<string, string>
        {
            { "LinkText_en", "Visit Aspose" },
            { "LinkText_es", "Visitar Aspose" }
        };

        // 2. Choose a culture (e.g., Spanish) and obtain the localized display text.
        string cultureCode = "es"; // Change to "en" for English.
        string resourceKey = $"LinkText_{cultureCode}";
        string localizedLinkText = resources.TryGetValue(resourceKey, out var txt) ? txt : "Default Link";

        // 3. Prepare the data model.
        var model = new ReportModel
        {
            Url = "https://www.aspose.com",
            LinkText = localizedLinkText
        };

        // 4. Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine("output", "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        // Insert a link tag where the first expression is the URL and the second is the display text.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");
        templateDoc.Save(templatePath);

        // 5. Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated document.
        string outputPath = Path.Combine("output", "Result.docx");
        doc.Save(outputPath);
    }
}
