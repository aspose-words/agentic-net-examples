using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

class Program
{
    static async Task Main()
    {
        // URLs of the source documents (any format supported by Aspose.Words)
        string urlFirst = "https://example.com/document1.docx";
        string urlSecond = "https://example.com/document2.docx";

        // Download the documents into memory streams
        using var httpClient = new HttpClient();
        using var streamFirst = new MemoryStream(await httpClient.GetByteArrayAsync(urlFirst));
        using var streamSecond = new MemoryStream(await httpClient.GetByteArrayAsync(urlSecond));

        // Load the documents (lifecycle rule)
        Document docFirst = new Document(streamFirst);
        Document docSecond = new Document(streamSecond);

        // Optional: configure comparison options
        CompareOptions compareOptions = new CompareOptions
        {
            Granularity = Granularity.WordLevel,
            Target = ComparisonTargetType.New
        };

        // Compare the documents – revisions are added to docFirst
        docFirst.Compare(docSecond, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result as PDF
        docFirst.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }
}
