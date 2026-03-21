#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class HyperlinkScanner
{
    // Entry point.
    static async Task Main()
    {
        // Path to the document that needs to be scanned.
        const string documentPath = @"C:\Docs\Sample.docx";

        // Load the document, or create a simple one if the file does not exist.
        Document doc;
        if (File.Exists(documentPath))
        {
            doc = new Document(documentPath);
        }
        else
        {
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document with a hyperlink.");
            // Insert a hyperlink field that points to a non‑existent file to demonstrate detection.
            builder.InsertField("HYPERLINK \"nonexistent.txt\"");
            // Insert a shape with an invalid web URL.
            Shape shape = new Shape(doc, ShapeType.Image);
            shape.HRef = "http://invalid.example.com/missing";
            builder.InsertNode(shape);
        }

        // Collect broken hyperlink information.
        List<string> brokenLinks = new List<string>();

        // -----------------------------------------------------------------
        // 1. Scan HYPERLINK fields (textual hyperlinks).
        // -----------------------------------------------------------------
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type != FieldType.FieldHyperlink)
                continue;

            // Cast to the strongly‑typed FieldHyperlink class.
            FieldHyperlink hyperlink = (FieldHyperlink)field;

            // The address may be a URL, a local file path, or a bookmark.
            string? address = hyperlink.Address?.Trim();
            if (string.IsNullOrEmpty(address))
                continue; // Nothing to validate.

            // If the hyperlink points to a bookmark (SubAddress is set), skip it – bookmarks are always valid.
            if (!string.IsNullOrEmpty(hyperlink.SubAddress))
                continue;

            // Determine whether the address is an external URL.
            bool isWebUrl = address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                            address.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

            if (isWebUrl)
            {
                // Perform a lightweight HEAD request to see if the URL is reachable.
                if (!await IsWebUrlReachable(address))
                    brokenLinks.Add($"Web URL not reachable: {address}");
            }
            else
            {
                // Treat the address as a file system path.
                // Resolve relative paths using the document's HyperlinkBase property if it is set.
                string basePath = doc.BuiltInDocumentProperties.HyperlinkBase;
                string combinedPath = string.IsNullOrEmpty(basePath)
                    ? Path.GetFullPath(Path.Combine(Path.GetDirectoryName(documentPath) ?? string.Empty, address))
                    : Path.GetFullPath(Path.Combine(basePath, address));

                if (!File.Exists(combinedPath))
                    brokenLinks.Add($"Missing file: {address} (resolved to {combinedPath})");
            }
        }

        // -----------------------------------------------------------------
        // 2. Scan shapes that contain hyperlinks (e.g., images with HRef).
        // -----------------------------------------------------------------
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (string.IsNullOrEmpty(shape.HRef))
                continue;

            string href = shape.HRef.Trim();

            // Same validation logic as for fields.
            bool isWebUrl = href.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                            href.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

            if (isWebUrl)
            {
                if (!await IsWebUrlReachable(href))
                    brokenLinks.Add($"Web URL in shape not reachable: {href}");
            }
            else
            {
                string basePath = doc.BuiltInDocumentProperties.HyperlinkBase;
                string combinedPath = string.IsNullOrEmpty(basePath)
                    ? Path.GetFullPath(Path.Combine(Path.GetDirectoryName(documentPath) ?? string.Empty, href))
                    : Path.GetFullPath(Path.Combine(basePath, href));

                if (!File.Exists(combinedPath))
                    brokenLinks.Add($"Missing file in shape: {href} (resolved to {combinedPath})");
            }
        }

        // -----------------------------------------------------------------
        // 3. Output the results.
        // -----------------------------------------------------------------
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("No broken hyperlinks were found.");
        }
        else
        {
            Console.WriteLine("Broken hyperlinks detected:");
            foreach (string entry in brokenLinks)
                Console.WriteLine("- " + entry);
        }

        // Save a simple text report to the temporary folder.
        string reportPath = Path.Combine(Path.GetTempPath(), "HyperlinkReport.txt");
        File.WriteAllLines(reportPath, brokenLinks);
        Console.WriteLine($"Report saved to: {reportPath}");
    }

    // Helper method that performs a HEAD request to check URL availability.
    private static async Task<bool> IsWebUrlReachable(string url)
    {
        try
        {
            using HttpClient client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            HttpResponseMessage response = await client.SendAsync(new HttpRequestMessage(HttpMethod.Head, url));
            return response.IsSuccessStatusCode;
        }
        catch
        {
            // Any exception (timeout, DNS failure, etc.) is treated as unreachable.
            return false;
        }
    }
}
