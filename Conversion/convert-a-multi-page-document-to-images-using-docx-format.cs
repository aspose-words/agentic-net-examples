using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the multi‑page DOCX document.
        Document doc = new Document("Input.docx");

        // Create image save options (PNG format in this example).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        // Optional: set desired resolution and image size.
        options.Resolution = 300;                     // 300 DPI
        options.ImageSize = new Size(1240, 1754);     // Approx. A4 at 300 DPI

        // Iterate through all pages and save each one as a separate image file.
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Render the current page (zero‑based index) only.
            options.PageSet = new PageSet(i);

            // Build the output file name.
            string outputFile = $"Page_{i + 1}.png";

            // Save the rendered page to an image file.
            doc.Save(outputFile, options);
        }
    }
}
