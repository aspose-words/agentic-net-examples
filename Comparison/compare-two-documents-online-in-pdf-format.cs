using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparison
{
    static void Main()
    {
        // URLs of the two source PDF documents to compare.
        const string urlA = "https://example.com/documents/sourceA.pdf";
        const string urlB = "https://example.com/documents/sourceB.pdf";

        // Download the documents into memory.
        using var httpClient = new HttpClient();

        byte[] bytesA = httpClient.GetByteArrayAsync(urlA).Result;
        byte[] bytesB = httpClient.GetByteArrayAsync(urlB).Result;

        // Load the PDFs into Aspose.Words Document objects.
        using var streamA = new MemoryStream(bytesA);
        using var streamB = new MemoryStream(bytesB);

        Document docA = new Document(streamA);
        Document docB = new Document(streamB);

        // Compare docA with docB. Revisions will be added to docA.
        docA.Compare(docB, "Comparer", DateTime.Now);

        // Save the comparison result as a PDF file.
        docA.Save("ComparedResult.pdf", SaveFormat.Pdf);
    }
}
