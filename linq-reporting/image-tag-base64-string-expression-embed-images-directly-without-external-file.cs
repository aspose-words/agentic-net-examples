using System;
using System.Drawing;
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

            // Insert an image from the local file system.
            // The image will be stored inside the document.
            builder.InsertImage(@"C:\Images\Sample.png");

            // Configure HTML save options to embed images as Base64 strings.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true, // Embed images directly in <img src="data:..."> tags.
                PrettyFormat = true          // Optional: make the output HTML more readable.
            };

            // Save the document as HTML with embedded Base64 images.
            doc.Save(@"C:\Output\DocumentWithBase64Images.html", saveOptions);
        }
    }
}
