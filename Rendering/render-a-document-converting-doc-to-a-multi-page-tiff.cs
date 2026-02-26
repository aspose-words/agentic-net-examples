using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToMultiPageTiff
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "input.doc";

        // Path where the multi‑page TIFF will be saved.
        string outputPath = "output.tiff";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Render each page as a separate frame in a multi‑frame TIFF.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Optional: set resolution (dpi) and compression.
        saveOptions.Resolution = 300;                     // 300 DPI
        saveOptions.TiffCompression = TiffCompression.Lzw; // LZW compression

        // Save the document as a multi‑page TIFF.
        doc.Save(outputPath, saveOptions);
    }
}
