using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToImageExample
{
    class Program
    {
        static void Main()
        {
            // Load the source Word document.
            Document doc = new Document("InputDocument.docx");

            // Create ImageSaveOptions for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Set custom resolution (dots per inch).
                Resolution = 300,

                // Adjust image brightness and contrast (0.0 – 1.0 range, default 0.5).
                ImageBrightness = 0.4f,
                ImageContrast = 0.6f,

                // Choose a color mode (None = full color, Grayscale, BlackAndWhite).
                ImageColorMode = ImageColorMode.Grayscale,

                // Set pixel format (e.g., 32‑bit ARGB).
                PixelFormat = ImagePixelFormat.Format32BppArgb,

                // Define a background (paper) color – transparent in this case.
                PaperColor = Color.Transparent,

                // Optionally scale the image (1.0 = 100%).
                Scale = 1.0f
            };

            // Save the first page of the document as an image using the custom options.
            doc.Save("OutputImage.png", options);
        }
    }
}
