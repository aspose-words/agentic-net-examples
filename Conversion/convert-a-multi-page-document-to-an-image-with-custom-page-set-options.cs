using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocumentToImage
{
    static void Main()
    {
        // Load the source Word document from disk.
        Document doc = new Document("Input.docx");

        // Create an ImageSaveOptions object for JPEG output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

        // Define a custom set of pages to render:
        // - Page 1 (zero‑based index 0)
        // - Pages 2‑3 (indices 1‑2)
        // - The last page of the document (using MaxValue as a shortcut)
        PageSet customPageSet = new PageSet(
            new PageRange(0, 0),               // page 1
            new PageRange(1, 2),               // pages 2‑3
            new PageRange(int.MaxValue, int.MaxValue) // last page
        );

        // Assign the custom page set to the save options.
        options.PageSet = customPageSet;

        // Optional: increase resolution and define a target image size.
        options.Resolution = 300; // 300 DPI
        options.ImageSize = new Size(2480, 3508); // Approx. A4 at 300 DPI

        // Save the selected pages as a single JPEG image.
        doc.Save("SelectedPages.jpg", options);
    }
}
