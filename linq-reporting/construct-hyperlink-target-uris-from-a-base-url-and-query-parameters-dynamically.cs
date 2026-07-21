using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HyperlinkLinqReportingExample
{
    // Model classes used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Base URL for the hyperlink.
        public string BaseUrl { get; set; } = string.Empty;

        // Collection of query parameters.
        public List<QueryParam> QueryParams { get; set; } = new();

        // Text that will be displayed for the hyperlink.
        public string LinkText { get; set; } = string.Empty;

        // Dynamically constructed full URL (base + encoded query string).
        public string FullUrl
        {
            get
            {
                if (string.IsNullOrEmpty(BaseUrl))
                    return string.Empty;

                if (QueryParams == null || QueryParams.Count == 0)
                    return BaseUrl;

                var encodedParams = QueryParams
                    .Select(p => $"{Uri.EscapeDataString(p.Name)}={Uri.EscapeDataString(p.Value)}");
                return $"{BaseUrl}?{string.Join("&", encodedParams)}";
            }
        }
    }

    public class QueryParam
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a blank document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Insert a LINQ Reporting link tag.
            //    The first expression provides the URI, the second provides the display text.
            builder.Writeln("<<link [model.FullUrl] [model.LinkText]>>");

            // 3. Prepare sample data.
            ReportModel model = new ReportModel
            {
                BaseUrl = "https://example.com/search",
                LinkText = "Search on Example.com",
                QueryParams = new List<QueryParam>
                {
                    new QueryParam { Name = "q", Value = "aspose words" },
                    new QueryParam { Name = "page", Value = "1" }
                }
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 5. Save the resulting document.
            doc.Save("HyperlinkReport.docx");
        }
    }
}
