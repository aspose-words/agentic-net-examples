using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string BaseUrl { get; set; } = "";
    public Dictionary<string, string> Parameters { get; set; } = new();
    public string LinkText { get; set; } = "";

    // Constructs the full URL with query parameters.
    public string FullUrl
    {
        get
        {
            if (Parameters == null || Parameters.Count == 0)
                return BaseUrl;

            var query = string.Join("&",
                Parameters.Select(kv => $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value)}"));
            return $"{BaseUrl}?{query}";
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            BaseUrl = "https://example.com/search",
            Parameters = new Dictionary<string, string>
            {
                { "q", "aspose" },
                { "page", "1" }
            },
            LinkText = "Search Aspose"
        };

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Dynamic Hyperlink Example:");
        // LINQ Reporting link tag using the computed FullUrl and LinkText.
        builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var loadedTemplate = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        loadedTemplate.Save(outputPath);
    }
}
