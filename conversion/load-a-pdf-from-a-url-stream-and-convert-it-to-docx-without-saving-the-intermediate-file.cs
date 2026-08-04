using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // URL of a sample PDF file.
        const string pdfUrl = "https://filesamples.com/samples/document/pdf/sample1.pdf";

        // Download the PDF into a byte array.
        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(pdfUrl).Result;
            response.EnsureSuccessStatusCode();
            byte[] pdfBytes = response.Content.ReadAsByteArrayAsync().Result;

            // Load the PDF from a memory stream.
            using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
            {
                Document pdfDocument = new Document(pdfStream);

                // Convert and save directly to DOCX.
                const string outputPath = "converted.docx";
                pdfDocument.Save(outputPath, SaveFormat.Docx);

                // Verify that the DOCX file was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("The DOCX output file was not created.");
            }
        }
    }
}
