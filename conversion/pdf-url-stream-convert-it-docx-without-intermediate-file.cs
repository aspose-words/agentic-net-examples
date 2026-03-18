using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // URL of the PDF document to be converted.
        const string pdfUrl = "https://example.com/sample.pdf";

        // Destination path for the resulting DOCX file.
        const string outputDocxPath = "ConvertedDocument.docx";

        // Download the PDF into a memory stream.
        using (HttpClient httpClient = new HttpClient())
        using (Stream pdfStream = httpClient.GetStreamAsync(pdfUrl).Result)
        using (MemoryStream pdfMemory = new MemoryStream())
        {
            // Copy the downloaded data to a seekable MemoryStream.
            pdfStream.CopyTo(pdfMemory);
            pdfMemory.Position = 0; // Reset position for reading.

            // Load the PDF directly from the memory stream.
            Document pdfDocument = new Document(pdfMemory);

            // Save the loaded document as DOCX without creating any intermediate files.
            pdfDocument.Save(outputDocxPath, SaveFormat.Docx);
        }

        Console.WriteLine($"PDF has been converted and saved to '{outputDocxPath}'.");
    }
}
