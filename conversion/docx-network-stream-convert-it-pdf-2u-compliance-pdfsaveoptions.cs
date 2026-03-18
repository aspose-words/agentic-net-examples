using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static async Task Main()
    {
        // URL of the DOCX file to download
        const string docxUrl = "https://example.com/sample.docx";

        // Download the DOCX into a seek‑able memory stream
        using (HttpClient httpClient = new HttpClient())
        using (Stream networkStream = await httpClient.GetStreamAsync(docxUrl))
        using (MemoryStream docStream = new MemoryStream())
        {
            await networkStream.CopyToAsync(docStream);
            docStream.Position = 0; // reset for reading

            // Load the document from the stream (lifecycle rule: Document(Stream))
            Document document = new Document(docStream);

            // Create PDF save options and set PDF/A‑2u compliance (rule: PdfSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA2u
            };

            // Save the document as PDF/A‑2u (lifecycle rule: Save(string, SaveOptions))
            document.Save("ConvertedDocument.pdf", pdfOptions);
        }
    }
}
