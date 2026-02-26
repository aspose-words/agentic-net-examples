using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentComparison
{
    static void Main()
    {
        // URLs of the two source documents (DOCX format recommended for comparison)
        const string urlOriginal = "https://filesamples.com/samples/document/docx/sample1.docx";
        const string urlEdited   = "https://filesamples.com/samples/document/docx/sample2.docx";

        // Download the first document into a MemoryStream and load it into an Aspose.Words Document.
        Document docOriginal = LoadDocumentFromUrl(urlOriginal);

        // Download the second document into a MemoryStream and load it into an Aspose.Words Document.
        Document docEdited = LoadDocumentFromUrl(urlEdited);

        // Ensure both documents have no revisions before comparison (required by Aspose.Words).
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the comparison result as a PDF file.
        docOriginal.Save("ComparisonResult.pdf", SaveFormat.Pdf);
    }

    // Helper method that downloads a document from a URL and returns an Aspose.Words Document.
    private static Document LoadDocumentFromUrl(string url)
    {
        using (HttpClient client = new HttpClient())
        {
            // Synchronously get the byte array of the file.
            byte[] data = client.GetByteArrayAsync(url).Result;

            // Load the document from the byte array using a MemoryStream.
            using (MemoryStream stream = new MemoryStream(data))
            {
                return new Document(stream);
            }
        }
    }
}
