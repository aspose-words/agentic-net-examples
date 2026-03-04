using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static async Task Main()
    {
        // URLs of the DOCX files to compare
        string urlOriginal = "https://example.com/document1.docx";
        string urlEdited   = "https://example.com/document2.docx";

        // Load both documents from the web
        Document docOriginal = await LoadDocumentFromUrlAsync(urlOriginal);
        Document docEdited   = await LoadDocumentFromUrlAsync(urlEdited);

        // Compare only if both documents have no existing revisions
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the comparison result (original document now contains revisions)
        docOriginal.Save("ComparedResult.docx");
    }

    // Helper method to download a DOCX file and load it into an Aspose.Words Document
    static async Task<Document> LoadDocumentFromUrlAsync(string url)
    {
        using HttpClient httpClient = new HttpClient();
        byte[] fileBytes = await httpClient.GetByteArrayAsync(url);
        using MemoryStream stream = new MemoryStream(fileBytes);
        return new Document(stream);
    }
}
