using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HyperlinkLinqReportingExample
{
    // Model representing the data for the report.
    public class ReportModel
    {
        // Base URL for the hyperlink.
        public string BaseUrl { get; set; } = string.Empty;

        // Collection of query parameters.
        public List<QueryParam> Params { get; set; } = new();

        // Text that will be displayed for the hyperlink.
        public string LinkText { get; set; } = string.Empty;

        // Dynamically constructed full URL with query parameters.
        public string FullUrl => BuildUrl();

        private string BuildUrl()
        {
            if (string.IsNullOrEmpty(BaseUrl))
                return string.Empty;

            var sb = new StringBuilder();
            foreach (var p in Params)
            {
                if (sb.Length > 0)
                    sb.Append('&');
                sb.Append($"{UrlEncoder.Default.Encode(p.Name)}={UrlEncoder.Default.Encode(p.Value)}");
            }

            var separator = BaseUrl.Contains("?") ? "&" : "?";
            return sb.Length > 0 ? $"{BaseUrl}{separator}{sb}" : BaseUrl;
        }
    }

    // Simple key/value pair representing a query parameter.
    public class QueryParam
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a hyperlink tag that will use the dynamically built URL.
            builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and prepare the data model.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel
            {
                BaseUrl = "https://example.com/search",
                LinkText = "Search Results",
                Params = new List<QueryParam>
                {
                    new QueryParam { Name = "q", Value = "Aspose.Words" },
                    new QueryParam { Name = "page", Value = "1" }
                }
            };

            // -----------------------------------------------------------------
            // Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(reportPath);

            // Output the result status.
            Console.WriteLine(success
                ? $"Report generated successfully: {Path.GetFullPath(reportPath)}"
                : "Report generation failed.");
        }
    }
}
