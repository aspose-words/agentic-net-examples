using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // URLs of the two DOCX files to compare.
        const string urlOriginal = "https://example.com/documents/original.docx";
        const string urlEdited   = "https://example.com/documents/edited.docx";

        // Download the documents into memory.
        byte[] originalBytes = DownloadDocument(urlOriginal);
        byte[] editedBytes   = DownloadDocument(urlEdited);

        // Load the documents from the byte arrays using the Document(Stream) constructor.
        Document docOriginal;
        Document docEdited;
        using (MemoryStream originalStream = new MemoryStream(originalBytes))
        using (MemoryStream editedStream   = new MemoryStream(editedBytes))
        {
            docOriginal = new Document(originalStream);
            docEdited   = new Document(editedStream);
        }

        // Ensure both documents have no revisions before comparison (required by Aspose.Words).
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. The result (revisions) will be stored in docOriginal.
            docOriginal.Compare(docEdited, "AI", DateTime.Now);
        }

        // Save the comparison result (original document with revisions) to a local file.
        docOriginal.Save("ComparisonResult.docx");
    }

    // Helper method to download a document from a URL and return its bytes.
    private static byte[] DownloadDocument(string url)
    {
        using (HttpClient client = new HttpClient())
        {
            HttpResponseMessage response = client.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();
            return response.Content.ReadAsByteArrayAsync().Result;
        }
    }
}
