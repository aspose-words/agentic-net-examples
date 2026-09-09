using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a sample data model.
        var model = new ReportModel
        {
            BaseUrl = "https://example.com/search",
            Params = new List<QueryParam>
            {
                new QueryParam { Name = "q", Value = "aspose" },
                new QueryParam { Name = "page", Value = "1" }
            },
            LinkText = "Search Aspose"
        };

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Insert a LINQ Reporting link tag that will be replaced with the constructed URI.
        builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

        // Populate the template using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model for the report.
public class ReportModel
{
    // Base URL part (e.g., https://example.com/search)
    public string BaseUrl { get; set; } = string.Empty;

    // Collection of query parameters.
    public List<QueryParam> Params { get; set; } = new();

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = string.Empty;

    // Full URL constructed from BaseUrl and Params.
    public string FullUrl => BuildUrl();

    // Helper method to build the complete URL with encoded query string.
    private string BuildUrl()
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return string.Empty;

        var query = string.Join("&",
            Params.Select(p => $"{Uri.EscapeDataString(p.Name)}={Uri.EscapeDataString(p.Value)}"));

        return string.IsNullOrEmpty(query) ? BaseUrl : $"{BaseUrl}?{query}";
    }
}

// Simple key/value pair representing a single query parameter.
public class QueryParam
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
