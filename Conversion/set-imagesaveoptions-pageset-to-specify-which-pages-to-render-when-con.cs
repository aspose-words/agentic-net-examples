using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create an ImageSaveOptions object for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Define the pages to render:
        // - Page 1 (zero‑based index 0) as a single page.
        // - Pages 2‑3 (zero‑based indices 1‑2) as a range.
        PageSet pageSet = new PageSet(
            new PageRange(0, 0),   // page 1
            new PageRange(1, 2)    // pages 2‑3
        );

        // Assign the PageSet to the options.
        options.PageSet = pageSet;

        // Save the document as a multi‑page TIFF containing only the specified pages.
        doc.Save("Output.tiff", options);
    }
}
