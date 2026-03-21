using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDynamicImage
{
    static async Task Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the image to be inserted.
        string imageUrl = "https://example.com/path/to/image.jpg";

        byte[] imageBytes;

        using (HttpClient http = new HttpClient())
        {
            try
            {
                imageBytes = await http.GetByteArrayAsync(imageUrl);
            }
            catch
            {
                // Fallback to a 1x1 transparent PNG if download fails.
                const string placeholderBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                imageBytes = Convert.FromBase64String(placeholderBase64);
            }
        }

        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            // Insert the image from the stream.
            Shape imageShape = builder.InsertImage(ms);
            // Optional: set hyperlink on the image (if desired).
            // imageShape.HRef = "https://example.com";
        }

        // Save the document to a file.
        doc.Save("DynamicImage.docx");
    }
}
