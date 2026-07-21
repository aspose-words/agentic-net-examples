using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // URL of the HTML page to be converted.
        const string url = "https://www.example.com";

        // Download the HTML content into a byte array.
        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();
            byte[] htmlBytes = response.Content.ReadAsByteArrayAsync().Result;

            // Load the HTML into an Aspose.Words Document via a memory stream.
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                Document doc = new Document(htmlStream);

                // Convert and save the document as PDF.
                const string outputPath = "output.pdf";
                doc.Save(outputPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException("The PDF file was not created.");
            }
        }
    }
}
