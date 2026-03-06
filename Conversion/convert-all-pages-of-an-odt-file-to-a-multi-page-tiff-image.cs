using System;
using Aspose.Words;
using Aspose.Words.Saving;

class OdtToMultiPageTiff
{
    static void Main()
    {
        // Path to the source ODT file.
        string inputPath = @"C:\Input\document.odt";

        // Path where the resulting multi‑page TIFF will be saved.
        string outputPath = @"C:\Output\document.tiff";

        // Load the ODT document.
        Document doc = new Document(inputPath);

        // Configure image save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Use the TiffFrames layout so each page becomes a separate frame
        // in the resulting multi‑frame TIFF image.
        options.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as a multi‑page TIFF.
        doc.Save(outputPath, options);
    }
}
