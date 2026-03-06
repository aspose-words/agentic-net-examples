using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create ImageSaveOptions for the desired image format (e.g., JPEG).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

        // Specify the resolution (dots per inch) for the rendered image.
        options.Resolution = 300; // Sets both horizontal and vertical DPI.

        // Define the page range to render.
        // The constructor PageSet(startPage, endPage) uses zero‑based page indices.
        // This example renders pages 2 through 4 (indices 1‑3).
        options.PageSet = new PageSet(1, 3);

        // Optional: adjust JPEG quality (0‑100). Higher values give better quality.
        options.JpegQuality = 90;

        // Save the selected pages as a single image file.
        doc.Save("Output.jpg", options);
    }
}
