using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            BaseUrl = "https://example.com/search",
            Params = new List<QueryParam>
            {
                new QueryParam { Name = "q", Value = "aspnet" },
                new QueryParam { Name = "page", Value = "1" }
            },
            LinkText = "Search Link"
        };

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting link tag that will be replaced with the constructed URI.
        // The tag uses the model's FullUrl property for the target and LinkText for display.
        builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HyperlinkReport.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// Root data model referenced by the template.
public class ReportModel
{
    // Base URL part (e.g., https://example.com/search)
    public string BaseUrl { get; set; } = string.Empty;

    // Collection of query parameters.
    public List<QueryParam> Params { get; set; } = new();

    // Text that will be displayed as the hyperlink.
    public string LinkText { get; set; } = string.Empty;

    // Dynamically constructed full URL with encoded query parameters.
    public string FullUrl => BuildUrl();

    private string BuildUrl()
    {
        if (string.IsNullOrEmpty(BaseUrl))
            return string.Empty;

        if (Params == null || Params.Count == 0)
            return BaseUrl;

        var query = string.Join("&",
            Params.Select(p => $"{WebUtility.UrlEncode(p.Name)}={WebUtility.UrlEncode(p.Value)}"));

        return $"{BaseUrl}?{query}";
    }
}

// Simple class representing a single query parameter.
public class QueryParam
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
