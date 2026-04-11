using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // URL of the HTML page to convert.
        const string url = "https://www.example.com/";

        // Determine the output PDF file path (in the current working directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedFromWeb.pdf");

        // Download the HTML content from the specified URL.
        byte[] htmlBytes;
        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(url).Result;
            response.EnsureSuccessStatusCode();
            htmlBytes = response.Content.ReadAsByteArrayAsync().Result;
        }

        // Load the downloaded HTML into an Aspose.Words Document using a memory stream.
        using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
        {
            Document document = new Document(htmlStream);
            // Convert and save the document as PDF.
            document.Save(outputPath, SaveFormat.Pdf);
        }

        // Validate that the PDF file was created and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("PDF conversion failed: the output file was not created or is empty.");
        }

        // Inform the user where the PDF was saved.
        Console.WriteLine($"PDF successfully saved to: {outputPath}");
    }
}
