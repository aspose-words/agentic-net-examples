using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            BaseUrl = "https://example.com/search",
            LinkText = "Search Results",
            QueryParams = new List<KeyValuePair<string, string>>
            {
                new("q", "aspose words"),
                new("page", "1"),
                new("lang", "en")
            }
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // Insert a LINQ Reporting link tag that will be replaced with a real hyperlink.
        builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyperlinkReport.docx");
        template.Save(outputPath);
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// Public data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Base URL part (e.g., https://example.com/search)
    public string BaseUrl { get; set; } = string.Empty;

    // Collection of query parameters.
    public List<KeyValuePair<string, string>> QueryParams { get; set; } = new();

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = string.Empty;

    // Dynamically constructed full URL with encoded query string.
    public string FullUrl
    {
        get
        {
            if (QueryParams == null || QueryParams.Count == 0)
                return BaseUrl;

            string query = string.Join("&",
                QueryParams.Select(p =>
                    $"{Uri.EscapeDataString(p.Key)}={Uri.EscapeDataString(p.Value)}"));
            return $"{BaseUrl}?{query}";
        }
    }
}
