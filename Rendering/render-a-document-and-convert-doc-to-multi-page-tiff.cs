using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToMultiPageTiff
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the multi‑page TIFF will be saved.
        string outputPath = @"C:\Docs\ResultDocument.tiff";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages into a single multi‑frame TIFF.
            PageLayout = MultiPageLayout.TiffFrames(),

            // Optional: set resolution (dpi) for the rendered images.
            Resolution = 300
        };

        // Save the document as a multi‑page TIFF.
        doc.Save(outputPath, saveOptions);
    }
}
