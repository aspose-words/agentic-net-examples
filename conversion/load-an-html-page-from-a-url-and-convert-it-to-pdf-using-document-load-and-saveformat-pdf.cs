using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // URL of the HTML page to convert.
        const string url = "https://www.example.com";

        // Download the HTML content.
        using (HttpClient httpClient = new HttpClient())
        {
            byte[] htmlBytes = httpClient.GetByteArrayAsync(url).Result;

            // Load the HTML into an Aspose.Words Document from a memory stream.
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                Document document = new Document(htmlStream);

                // Convert and save the document as PDF.
                const string outputPath = "output.pdf";
                document.Save(outputPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(outputPath))
                {
                    throw new InvalidOperationException("The PDF file was not created.");
                }
            }
        }
    }
}
