using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Loading;

class HtmlToPdfConverter
{
    static void Main()
    {
        // URL of the HTML page to convert.
        const string url = "https://example.com/sample.html";

        // Destination PDF file path.
        const string outputPdfPath = "ConvertedDocument.pdf";

        // Download the HTML content as a byte array.
        using (HttpClient httpClient = new HttpClient())
        {
            byte[] htmlBytes = httpClient.GetByteArrayAsync(url).GetAwaiter().GetResult();

            // Load the HTML into a memory stream.
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Configure load options: specify that the format is HTML and provide the base URI
                // so that relative resources (e.g., images) are resolved correctly.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html, "", url);

                // Load the HTML document from the stream using the specified options.
                Document document = new Document(htmlStream, loadOptions);

                // Save the loaded document as PDF.
                document.Save(outputPdfPath, SaveFormat.Pdf);
            }
        }
    }
}
