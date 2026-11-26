using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result documents.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // ---------- Create template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a hyperlink tag that will be populated at runtime.
        // The tag uses the model's FullUrl property for the target and LinkText for display.
        builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load template ----------
        Document doc = new Document(templatePath);

        // ---------- Prepare data model ----------
        ReportModel model = new()
        {
            BaseUrl = "https://example.com/search",
            Parameters = new()
            {
                new Param { Name = "q", Value = "aspose words" },
                new Param { Name = "page", Value = "1" }
            },
            LinkText = "Search on Example.com"
        };

        // ---------- Build report ----------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(resultPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    // Base URL without query string.
    public string BaseUrl { get; set; } = string.Empty;

    // Collection of query parameters.
    public List<Param> Parameters { get; set; } = new();

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = string.Empty;

    // Dynamically constructed full URL with encoded query parameters.
    public string FullUrl
    {
        get
        {
            if (Parameters == null || Parameters.Count == 0)
                return BaseUrl;

            string query = string.Join("&",
                Parameters.Select(p =>
                    $"{WebUtility.UrlEncode(p.Name)}={WebUtility.UrlEncode(p.Value)}"));
            return $"{BaseUrl}?{query}";
        }
    }
}

// Simple key‑value pair representing a query parameter.
public class Param
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
