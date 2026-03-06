using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToMultiPageTiff
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.doc");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Use the TiffFrames layout to render each page as a separate frame
        // in a single multi‑page TIFF file.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: set compression (default is Lzw).
        // saveOptions.TiffCompression = TiffCompression.Lzw;

        // Save the document as a multi‑page TIFF.
        doc.Save("output.tiff", saveOptions);
    }
}
