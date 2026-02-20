using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    // Entry point.
    static async Task Main()
    {
        // URLs of the two DOCX documents to compare.
        const string urlDocA = "https://example.com/documentA.docx";
        const string urlDocB = "https://example.com/documentB.docx";

        // Download the documents into memory streams.
        using var httpClient = new HttpClient();

        using var streamA = new MemoryStream(await httpClient.GetByteArrayAsync(urlDocA));
        using var streamB = new MemoryStream(await httpClient.GetByteArrayAsync(urlDocB));

        // Load the documents from the streams.
        Document docA = new Document(streamA);
        Document docB = new Document(streamB);

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: track changes at the word level and compare moves.
            Granularity = Granularity.WordLevel,
            CompareMoves = true,
            // Adjust other flags as needed.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. The revisions will be added to docA.
        docA.Compare(docB, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document with revisions.
        const string outputPath = "ComparedResult.docx";
        docA.Save(outputPath);
    }
}
