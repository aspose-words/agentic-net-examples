using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source RTF document (multi‑page).
        string rtfPath = "input.rtf";

        // Path where the resulting multi‑frame TIFF will be saved.
        string tiffPath = "output.tiff";

        // Load the RTF document.
        Document doc = new Document(rtfPath);

        // Create image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Configure the layout so that each page is rendered as a separate frame
        // in a single multi‑page TIFF image.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as a multi‑page TIFF.
        doc.Save(tiffPath, saveOptions);
    }
}
