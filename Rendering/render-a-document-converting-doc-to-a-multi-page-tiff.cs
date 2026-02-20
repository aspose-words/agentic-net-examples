using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Render each page as a separate frame in a multi‑page TIFF.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: choose a compression scheme (e.g., LZW).
        saveOptions.TiffCompression = TiffCompression.Lzw;

        // Save the document as a multi‑page TIFF file.
        doc.Save("Output.tiff", saveOptions);
    }
}
