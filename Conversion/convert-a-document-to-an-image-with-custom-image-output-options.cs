using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set resolution (dots per inch).
            Resolution = 300,

            // Adjust image brightness and contrast (0.0 to 1.0, default 0.5).
            ImageBrightness = 0.4f,
            ImageContrast = 0.6f,

            // Set horizontal and vertical resolution; Scale compensates size change.
            HorizontalResolution = 72f,
            VerticalResolution = 72f,
            Scale = 96f / 72f,

            // Use a transparent background.
            PaperColor = Color.Transparent,

            // Define explicit pixel dimensions.
            ImageSize = new Size(1200, 1600),

            // Choose pixel format (32‑bit ARGB).
            PixelFormat = ImagePixelFormat.Format32BppArgb,

            // Render the image in grayscale.
            ImageColorMode = ImageColorMode.Grayscale
        };

        // Save the first page of the document as an image with the specified options.
        doc.Save("Output.png", options);
    }
}
