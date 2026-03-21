using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsBase64Example
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Use DocumentBuilder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A 1x1 pixel transparent PNG encoded as Base64.
            const string pngBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(pngBase64);

            // Insert the image from the byte array.
            builder.InsertImage(pngBytes);

            // Configure HTML save options to embed images as Base64 strings.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true, // Embed images directly in <img src="data:..."> tags.
                PrettyFormat = true          // Optional: make the output HTML more readable.
            };

            // Save the document as HTML with embedded Base64 images in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithBase64Images.html");
            doc.Save(outputPath, saveOptions);

            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
