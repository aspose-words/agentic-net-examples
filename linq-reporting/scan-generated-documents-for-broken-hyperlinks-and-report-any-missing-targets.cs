using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Fields;

public class HyperlinkScanner
{
    // Entry point of the console application.
    public static void Main()
    {
        // Create a sample document containing various hyperlinks.
        string docPath = "Sample.docx";
        CreateSampleDocument(docPath);

        // Load the document for scanning.
        Document doc = new Document(docPath);

        // Scan the document for broken hyperlinks.
        List<string> brokenLinks = ScanDocumentForBrokenHyperlinks(doc);

        // Report the results.
        Console.WriteLine("Hyperlink scan completed.");
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("No broken hyperlinks were found.");
        }
        else
        {
            Console.WriteLine("Broken hyperlinks:");
            foreach (string link in brokenLinks)
                Console.WriteLine("- " + link);
        }
    }

    // Creates a Word document with a mix of valid and invalid hyperlinks.
    private static void CreateSampleDocument(string fileName)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Valid local file (will be created).
        string validLocalFile = "ExistingFile.txt";
        File.WriteAllText(validLocalFile, "Sample content");
        builder.Write("Valid local file: ");
        builder.InsertHyperlink("Open existing file", validLocalFile, false);
        builder.Writeln();

        // Invalid local file (does not exist).
        string missingLocalFile = "MissingFile.txt";
        builder.Write("Missing local file: ");
        builder.InsertHyperlink("Open missing file", missingLocalFile, false);
        builder.Writeln();

        // Valid URL.
        string validUrl = "https://www.example.com/";
        builder.Write("Valid URL: ");
        builder.InsertHyperlink("Visit Example.com", validUrl, false);
        builder.Writeln();

        // Invalid URL (expected to fail).
        string invalidUrl = "https://nonexistent.example.invalid/";
        builder.Write("Invalid URL: ");
        builder.InsertHyperlink("Visit broken link", invalidUrl, false);
        builder.Writeln();

        // Bookmark within the document (always valid).
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Bookmark target text.");
        builder.EndBookmark("MyBookmark");
        builder.Write("Internal bookmark: ");
        builder.InsertHyperlink("Go to bookmark", "MyBookmark", true);
        builder.Writeln();

        // Save the document.
        doc.Save(fileName);
    }

    // Scans a document for hyperlinks whose targets cannot be resolved.
    private static List<string> ScanDocumentForBrokenHyperlinks(Document doc)
    {
        var brokenLinks = new List<string>();
        // HttpClient is intended to be reused; dispose at the end.
        using var httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(5)
        };

        // Iterate over all fields in the document.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type != FieldType.FieldHyperlink)
                continue;

            var hyperlink = (FieldHyperlink)field;
            string address = hyperlink.Address?.Trim() ?? string.Empty;
            string subAddress = hyperlink.SubAddress?.Trim() ?? string.Empty;

            // If the hyperlink points to a bookmark inside the document, consider it valid.
            if (string.IsNullOrEmpty(address) && !string.IsNullOrEmpty(subAddress))
                continue;

            // Empty address means nothing to check.
            if (string.IsNullOrEmpty(address))
                continue;

            // Determine if the address is a URL or a file path.
            if (address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                address.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                // Attempt a HEAD request to verify the URL is reachable.
                try
                {
                    using var request = new HttpRequestMessage(HttpMethod.Head, address);
                    using var response = httpClient.SendAsync(request).Result;
                    if (!response.IsSuccessStatusCode)
                        brokenLinks.Add(address);
                }
                catch
                {
                    // Any exception indicates the URL is not reachable.
                    brokenLinks.Add(address);
                }
            }
            else
            {
                // Resolve relative paths against the current directory.
                string resolvedPath = Path.GetFullPath(address);
                if (!File.Exists(resolvedPath))
                    brokenLinks.Add(address);
            }
        }

        return brokenLinks;
    }
}
