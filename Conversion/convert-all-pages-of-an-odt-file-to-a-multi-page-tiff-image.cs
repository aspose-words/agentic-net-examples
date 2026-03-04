using System;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class OdtToMultiPageTiff
{
    static void Main()
    {
        // Path to the source ODT file.
        string inputPath = @"C:\Input\sample.odt";

        // Path where the resulting multi‑page TIFF will be saved.
        string outputPath = @"C:\Output\sample.tiff";

        // Load the ODT document.
        Document doc = new Document(inputPath);

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages into a single multi‑frame TIFF.
            PageLayout = MultiPageLayout.TiffFrames(),

            // Optional: set resolution (dpi) for higher quality.
            Resolution = 300,

            // Optional: set background color if desired (default is white).
            PaperColor = Color.White
        };

        // Save the document as a multi‑page TIFF.
        doc.Save(outputPath, saveOptions);
    }
}
