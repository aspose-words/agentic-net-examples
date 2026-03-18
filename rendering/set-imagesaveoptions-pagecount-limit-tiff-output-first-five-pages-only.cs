using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Limit the output to the first five pages (zero‑based indices 0‑4).
        tiffOptions.PageSet = new PageSet(0, 1, 2, 3, 4);

        // Save the document as a multi‑page TIFF containing only those pages.
        doc.Save("Output.tiff", tiffOptions);
    }
}
