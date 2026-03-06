using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // URL of the sample DOCX file to download.
        const string sampleUrl = "https://example.com/sample.docx";

        // Destination path for the converted PDF file.
        const string outputPdfPath = "ConvertedDocument.pdf";

        // Download the DOCX file into a byte array.
        byte[] docxBytes;
        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(sampleUrl).Result;
            response.EnsureSuccessStatusCode();
            docxBytes = response.Content.ReadAsByteArrayAsync().Result;
        }

        // Load the downloaded DOCX from a memory stream using Aspose.Words.
        using (MemoryStream docxStream = new MemoryStream(docxBytes))
        {
            Document doc = new Document(docxStream); // Load constructor (Document(Stream))

            // Save the document as PDF. The format is inferred from the .pdf extension.
            doc.Save(outputPdfPath); // Save(string) – automatic format detection
        }

        Console.WriteLine($"Document has been converted and saved to '{outputPdfPath}'.");
    }
}
