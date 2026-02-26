using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // URLs of the two DOCX files to compare.
        string urlOriginal = "https://example.com/documents/original.docx";
        string urlEdited   = "https://example.com/documents/edited.docx";

        // Load both documents from the web into Aspose.Words Document objects.
        Document docOriginal = LoadDocumentFromUrl(urlOriginal);
        Document docEdited   = LoadDocumentFromUrl(urlEdited);

        // Ensure both documents have no existing revisions before comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Compare the documents. Revisions are added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now);
        }

        // Save the result (original document with revisions) to a local file.
        docOriginal.Save("ComparisonResult.docx");
    }

    // Downloads a DOCX file from the specified URL and loads it into a Document.
    private static Document LoadDocumentFromUrl(string url)
    {
        using (HttpClient client = new HttpClient())
        {
            // Retrieve the file bytes.
            byte[] data = client.GetByteArrayAsync(url).Result;

            // Load the document from a memory stream.
            using (MemoryStream stream = new MemoryStream(data))
            {
                return new Document(stream);
            }
        }
    }
}
