using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Fields;

namespace HyperlinkScanner
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            // Create a sample document with various hyperlinks.
            string docPath = "GeneratedDocument.docx";
            CreateSampleDocument(docPath);

            // Load the document for scanning.
            Document doc = new Document(docPath);

            // Scan hyperlinks and report broken ones.
            await ScanHyperlinksAsync(doc);
        }

        private static void CreateSampleDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Valid external URL.
            builder.Writeln("Valid URL:");
            builder.InsertHyperlink("Google", "https://www.google.com", false);
            builder.Writeln();

            // Invalid external URL.
            builder.Writeln("Invalid URL:");
            builder.InsertHyperlink("BrokenLink", "https://nonexistent.example.com", false);
            builder.Writeln();

            // Missing local file.
            builder.Writeln("Missing local file:");
            builder.InsertHyperlink("MissingFile", @"C:\nonexistent\file.txt", false);
            builder.Writeln();

            // Bookmark target.
            builder.StartBookmark("MyBookmark");
            builder.Writeln("This is the bookmark target.");
            builder.EndBookmark("MyBookmark");

            // Link to existing bookmark.
            builder.Writeln("Link to bookmark:");
            builder.InsertHyperlink("GoToBookmark", "MyBookmark", true);
            builder.Writeln();

            // Link to non‑existent bookmark.
            builder.Writeln("Link to missing bookmark:");
            builder.InsertHyperlink("MissingBookmark", "NoSuchBookmark", true);
            builder.Writeln();

            doc.Save(filePath);
        }

        private static async Task ScanHyperlinksAsync(Document doc)
        {
            using HttpClient httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(5)
            };

            var fields = doc.Range.Fields
                .OfType<FieldHyperlink>()
                .ToList();

            foreach (var hyperlink in fields)
            {
                bool isBroken = await IsHyperlinkBrokenAsync(hyperlink, doc, httpClient);
                string displayText = hyperlink.Result?.Trim() ?? "(no text)";
                string target = !string.IsNullOrEmpty(hyperlink.Address)
                    ? hyperlink.Address
                    : hyperlink.SubAddress ?? "(no target)";

                Console.WriteLine($"Hyperlink \"{displayText}\" -> \"{target}\": {(isBroken ? "BROKEN" : "OK")}");
            }
        }

        private static async Task<bool> IsHyperlinkBrokenAsync(FieldHyperlink hyperlink, Document doc, HttpClient httpClient)
        {
            // Check external address (URL or file path).
            if (!string.IsNullOrEmpty(hyperlink.Address))
            {
                string address = hyperlink.Address;

                // Treat HTTP/HTTPS as web URLs.
                if (address.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        using var response = await httpClient.GetAsync(address);
                        return !response.IsSuccessStatusCode;
                    }
                    catch
                    {
                        return true; // Any exception means the link is unreachable.
                    }
                }
                else
                {
                    // Assume a file system path.
                    string path = address;
                    if (!Path.IsPathRooted(path))
                    {
                        path = Path.Combine(Directory.GetCurrentDirectory(), path);
                    }
                    return !File.Exists(path);
                }
            }

            // Check bookmark reference.
            if (!string.IsNullOrEmpty(hyperlink.SubAddress))
            {
                return !doc.Range.Bookmarks.Any(b => b.Name.Equals(hyperlink.SubAddress, StringComparison.Ordinal));
            }

            // No address or bookmark – consider it broken.
            return true;
        }
    }
}
