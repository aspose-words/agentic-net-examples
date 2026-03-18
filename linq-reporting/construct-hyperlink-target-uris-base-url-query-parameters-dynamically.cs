using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Aspose.Words;
using Aspose.Words.Fields;

class HyperlinkBuilder
{
    static void Main()
    {
        // Base URL for the hyperlink
        string baseUrl = "https://example.com/search";

        // Query parameters to be appended dynamically
        var queryParams = new Dictionary<string, string>
        {
            { "q", "aspose words" },
            { "page", "1" },
            { "lang", "en" }
        };

        // Build the full URI with encoded query string
        string fullUri = BuildUri(baseUrl, queryParams);

        // Create a new document and a builder (lifecycle rule)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the hyperlink with display text
        builder.Font.Color = System.Drawing.Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Search Aspose Words", fullUri, false);
        builder.Font.ClearFormatting();

        // Save the document (lifecycle rule)
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DynamicHyperlink.docx");
        doc.Save(outputPath);
    }

    // Helper method to construct a URI from base and parameters
    private static string BuildUri(string baseUrl, IDictionary<string, string> parameters)
    {
        if (parameters == null || parameters.Count == 0)
            return baseUrl;

        var encodedParams = parameters
            .Select(kv => $"{WebUtility.UrlEncode(kv.Key)}={WebUtility.UrlEncode(kv.Value)}");

        string queryString = string.Join("&", encodedParams);
        return $"{baseUrl}?{queryString}";
    }
}
