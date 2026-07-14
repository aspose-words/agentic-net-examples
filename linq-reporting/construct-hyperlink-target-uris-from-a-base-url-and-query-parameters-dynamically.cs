using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Base URL for the hyperlink.
    public string BaseUrl { get; set; } = string.Empty;

    // Collection of query parameters.
    public Dictionary<string, string> Parameters { get; set; } = new();

    // Text that will be displayed for the hyperlink.
    public string LinkText { get; set; } = string.Empty;

    // Dynamically constructed full URL with encoded query string.
    public string FullUrl
    {
        get
        {
            if (Parameters == null || Parameters.Count == 0)
                return BaseUrl;

            var query = string.Join("&",
                Parameters.Select(p =>
                    $"{Uri.EscapeDataString(p.Key)}={Uri.EscapeDataString(p.Value)}"));
            return $"{BaseUrl}?{query}";
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Dynamic Hyperlink Example");
        // LINQ Reporting link tag: <<link [uriExpression] [displayTextExpression]>>
        builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        var model = new ReportModel
        {
            BaseUrl = "https://example.com/search",
            Parameters = new Dictionary<string, string>
            {
                { "q", "Aspose.Words" },
                { "page", "1" },
                { "lang", "en" }
            },
            LinkText = "Search Aspose.Words"
        };

        // -----------------------------------------------------------------
        // 3. Build the report using Aspose.Words LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
