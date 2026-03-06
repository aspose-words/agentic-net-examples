using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DownsampleDocument
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("MyDir/Images.docx");

        // Create PDF save options that include downsample settings.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Enable image downsampling.
        saveOptions.DownsampleOptions.DownsampleImages = true;

        // Set the target resolution for downsampled images (e.g., 72 DPI).
        saveOptions.DownsampleOptions.Resolution = 72;

        // No threshold – downsample all images that can be reduced.
        saveOptions.DownsampleOptions.ResolutionThreshold = 0;

        // Save the document as PDF with the configured downsampling.
        doc.Save("ArtifactsDir/Downsampled.pdf", saveOptions);
    }
}
