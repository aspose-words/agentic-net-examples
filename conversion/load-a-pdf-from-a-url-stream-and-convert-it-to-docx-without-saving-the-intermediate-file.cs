using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // URL of a sample PDF file.
        const string pdfUrl = "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf";

        // Path for the resulting DOCX file.
        const string outputPath = "output.docx";

        // Download the PDF into a memory stream.
        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(pdfUrl).Result;
            response.EnsureSuccessStatusCode();
            byte[] pdfBytes = response.Content.ReadAsByteArrayAsync().Result;

            using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
            {
                // Ensure the stream is positioned at the beginning before loading.
                pdfStream.Position = 0;

                // Load the PDF document directly from the stream.
                Document pdfDocument = new Document(pdfStream);

                // Convert and save the document as DOCX without creating an intermediate file.
                pdfDocument.Save(outputPath, SaveFormat.Docx);
            }
        }

        // Verify that the DOCX file was created and contains data.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("Conversion failed: DOCX file was not created or is empty.");
        }
    }
}
