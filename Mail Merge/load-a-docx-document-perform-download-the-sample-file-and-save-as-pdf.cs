using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // URL of the sample DOCX file to download
        const string docxUrl = "https://example.com/sample.docx";

        // Path where the resulting PDF will be saved
        const string pdfPath = "sample.pdf";

        // Download the DOCX file into a byte array
        using (HttpClient httpClient = new HttpClient())
        {
            byte[] docxBytes = httpClient.GetByteArrayAsync(docxUrl).Result;

            // Load the downloaded DOCX from a memory stream using Aspose.Words Document(Stream) constructor
            using (MemoryStream docxStream = new MemoryStream(docxBytes))
            {
                Document document = new Document(docxStream);

                // Save the document as PDF; the format is inferred from the .pdf extension
                document.Save(pdfPath);
            }
        }
    }
}
