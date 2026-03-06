using Aspose.Words;
using Aspose.Words.Saving;

class RtfToMultiPageTiff
{
    static void Main()
    {
        // Path to the source RTF document.
        const string inputPath = @"C:\Docs\source.rtf";

        // Path where the resulting multi‑page TIFF will be saved.
        const string outputPath = @"C:\Docs\result.tiff";

        // Load the RTF document.
        Document doc = new Document(inputPath);

        // Create image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Configure the layout so that each page becomes a separate frame
        // in a single multi‑frame TIFF image.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as a multi‑page TIFF.
        doc.Save(outputPath, saveOptions);
    }
}
